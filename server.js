import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import unzipper from "unzipper";
import crypto from "crypto";
import { spawn, execFileSync } from "child_process";
import { XMLParser } from "fast-xml-parser";
import { MathMLToLaTeX } from "mathml-to-latex";

/* ================== FAST SETTINGS ================== */
const DEBUG = process.env.DEBUG_CONVERT === "1";
const MAX_OLE_CONCURRENCY = Number(process.env.OLE_CONCURRENCY || 6);   // ruby jobs in-flight
const MAX_IMG_CONCURRENCY = Number(process.env.IMG_CONCURRENCY || 2);   // inkscape jobs in-flight
const CACHE_OLE_MAX = Number(process.env.CACHE_OLE_MAX || 4000);
const CACHE_IMG_MAX = Number(process.env.CACHE_IMG_MAX || 1200);
const INKSCAPE_BIN = process.env.INKSCAPE_BIN || "inkscape";

// Detect sqrt in MathML
const SQRT_MATHML_RE = /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

/* ================== APP ================== */
const app = express();
app.use(cors());
app.use(express.json({ limit: "25mb" }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/* ================== UTIL ================== */
function safeUnlink(p) { try { fs.unlinkSync(p); } catch {} }
function safeRmdir(dir) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch {} }
function sha1(buf) { return crypto.createHash("sha1").update(buf).digest("hex"); }

/** tiny LRU */
class LRU {
  constructor(max = 1000) { this.max = max; this.map = new Map(); }
  get(k) {
    if (!this.map.has(k)) return undefined;
    const v = this.map.get(k);
    this.map.delete(k);
    this.map.set(k, v);
    return v;
  }
  set(k, v) {
    if (this.map.has(k)) this.map.delete(k);
    this.map.set(k, v);
    if (this.map.size > this.max) {
      const first = this.map.keys().next().value;
      this.map.delete(first);
    }
  }
}

/** Semaphore limiter */
class Semaphore {
  constructor(max) { this.max = max; this.cur = 0; this.q = []; }
  async acquire() {
    if (this.cur < this.max) { this.cur++; return; }
    await new Promise((res) => this.q.push(res));
    this.cur++;
  }
  release() {
    this.cur--;
    const next = this.q.shift();
    if (next) next();
  }
  async run(fn) {
    await this.acquire();
    try { return await fn(); }
    finally { this.release(); }
  }
}

async function openDocxZip(docxBuffer) { return unzipper.Open.buffer(docxBuffer); }
async function readZipEntry(zip, p) {
  const f = zip.__fileByPath?.get(p) || (zip.files || []).find((x) => x.path === p);
  if (!f) return null;
  return await f.buffer();
}
function unique(arr) { return [...new Set(arr || [])].filter(Boolean); }

function mimeFromExt(p) {
  const ext = (p.split(".").pop() || "").toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "emf") return "image/emf";
  if (ext === "wmf") return "image/wmf";
  return "application/octet-stream";
}
function getExtFromPath(p) { return (p.split(".").pop() || "").toLowerCase(); }

/* ================== RUBY WORKER (PERSISTENT) ================== */
class RubyMtWorker {
  constructor() {
    this.proc = null;
    this.seq = 0;
    this.pending = new Map();
    this.buf = "";
  }
  start() {
    if (this.proc) return;
    const script = path.join(process.cwd(), "mt2mml_worker.rb");
    if (!fs.existsSync(script)) {
      throw new Error("Missing mt2mml_worker.rb in cwd");
    }

    this.proc = spawn("ruby", [script], { stdio: ["pipe", "pipe", "pipe"] });

    this.proc.stdout.setEncoding("utf8");
    this.proc.stderr.setEncoding("utf8");

    this.proc.stdout.on("data", (chunk) => {
      this.buf += chunk;
      let idx;
      while ((idx = this.buf.indexOf("\n")) >= 0) {
        const line = this.buf.slice(0, idx).trim();
        this.buf = this.buf.slice(idx + 1);
        if (!line) continue;

        let msg;
        try { msg = JSON.parse(line); } catch { continue; }
        const id = msg.id;
        const p = this.pending.get(id);
        if (!p) continue;
        this.pending.delete(id);

        if (msg.ok) p.resolve(msg.mathml || "");
        else p.resolve("");
      }
    });

    this.proc.stderr.on("data", (chunk) => {
      if (DEBUG) console.warn("[RUBY_WORKER]", chunk.trim());
    });

    this.proc.on("exit", (code) => {
      if (DEBUG) console.warn("[RUBY_WORKER_EXIT]", code);
      this.proc = null;
      // fail all pending
      for (const [, p] of this.pending) p.resolve("");
      this.pending.clear();
    });
  }

  async convertOleBufferToMathML(oleBuf) {
    this.start();
    const id = `${Date.now()}_${(++this.seq)}_${Math.random().toString(16).slice(2)}`;
    const b64 = oleBuf.toString("base64");
    const payload = JSON.stringify({ id, b64, prefer_v2: true }) + "\n";

    return await new Promise((resolve) => {
      this.pending.set(id, { resolve });
      this.proc.stdin.write(payload);
    });
  }
}

const rubyWorker = new RubyMtWorker();
const oleLimiter = new Semaphore(MAX_OLE_CONCURRENCY);
const imgLimiter = new Semaphore(MAX_IMG_CONCURRENCY);

// caches
const oleCache = new LRU(CACHE_OLE_MAX); // key sha1(ole) -> {mathml, latex}
const imgCache = new LRU(CACHE_IMG_MAX); // key sha1(media+ext) -> dataUri

/* ================== INKSCAPE EMF/WMF -> PNG ================== */
function convertEmfWmfToPngInkscape(buffer, ext) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "img-convert-"));
  const inFile = path.join(tmpDir, `input.${ext}`);
  const outFile = path.join(tmpDir, `output.png`);

  try {
    fs.writeFileSync(inFile, buffer);

    // inkscape input.emf --export-type=png --export-filename=output.png
    execFileSync(
      INKSCAPE_BIN,
      [inFile, "--export-type=png", `--export-filename=${outFile}`],
      { stdio: "ignore", timeout: 30000 }
    );

    if (fs.existsSync(outFile)) return fs.readFileSync(outFile);
    return null;
  } catch (e) {
    if (DEBUG) console.warn("[INKSCAPE_FAIL]", ext, e?.message || String(e));
    return null;
  } finally {
    safeRmdir(tmpDir);
  }
}

/* ================== MATHML -> LATEX (keep your logic) ================== */
function preprocessMathMLForSqrt(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);

  const moSqrt = String.raw`<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>`;
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mrow>([\s\S]*?)<\/mrow>`, "gi"), "<msqrt>$1</msqrt>");
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mi>([^<]+)<\/mi>`, "gi"), "<msqrt><mi>$1</mi></msqrt>");
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mn>([^<]+)<\/mn>`, "gi"), "<msqrt><mn>$1</mn></msqrt>");
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mfenced([^>]*)>([\s\S]*?)<\/mfenced>`, "gi"), "<msqrt><mfenced$1>$2</mfenced></msqrt>");
  return s;
}
function postprocessLatexSqrt(latex) {
  if (!latex) return latex;
  let s = String(latex);
  s = s.replace(/\\surd\b/g, "\\sqrt{}");
  s = s.replace(/√\s*\{([^}]+)\}/g, "\\sqrt{$1}");
  s = s.replace(/√\s*\(([^)]+)\)/g, "\\sqrt{$1}");
  s = s.replace(/√\s*(\d+)/g, "\\sqrt{$1}");
  s = s.replace(/√\s*([a-zA-Z])/g, "\\sqrt{$1}");
  s = s.replace(/\\sqrt\s+(\d+)(?![}\d])/g, "\\sqrt{$1}");
  s = s.replace(/\\sqrt\s+([a-zA-Z])(?![}\w])/g, "\\sqrt{$1}");
  s = s.replace(/\\sqrt\s*\{\s*\}/g, "\\sqrt{\\phantom{x}}");
  s = s.replace(/\\sqrt\s+\{/g, "\\sqrt{");
  s = s.replace(/\\root\s*\{([^}]+)\}\s*\\of\s*\{([^}]+)\}/g, "\\sqrt[$1]{$2}");
  s = s.replace(/\\sqrt\s*\[\s*(\d+)\s*\]\s*\{/g, "\\sqrt[$1]{");
  return s;
}
function finalLatexCleanup(latex) {
  if (!latex) return latex;
  let s = String(latex);
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, "");
  s = s.replace(/[\u00A0]/g, " ");
  s = s.replace(/[\u2000-\u200A\u202F\u205F\u3000]/g, " ");
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
  s = s.replace(/\\left\s*\(\s*\*\s*\\right\s*\)/g, "(*)");
  s = s.replace(/\\left\s*\(\s*\\star\s*\\right\s*\)/g, "(*)");
  s = s.replace(/\\left\s*\(\s*\\right\s*\./g, "(");
  s = s.replace(/\\left\s*\.\s*\\right\s*\)/g, ")");
  s = s.replace(/\\left\s*\(\s*\\right\s*\)/g, "()");
  s = s.replace(/\bl\s+o\s+g\b/gi, "\\log");
  s = s.replace(/\bs\s+i\s+n\b/gi, "\\sin");
  s = s.replace(/\bc\s+o\s+s\b/gi, "\\cos");
  s = s.replace(/\bt\s+a\s+n\b/gi, "\\tan");
  s = s.replace(/\bl\s+n\b/gi, "\\ln");
  s = s.replace(/\bl\s+i\s+m\b/gi, "\\lim");
  s = s.replace(/\\log\s*(\d+)\s*_\s*\{\s*\}/g, "\\log_{$1}");
  s = s.replace(/\\log\s+(\d+)\s*\(/g, "\\log_{$1}(");
  s = s.replace(/\\log\s+(\d+)\s*\\left/g, "\\log_{$1}\\left");
  s = s.replace(/_\s*\{\s*\}/g, "");
  s = s.replace(/\^\s*\{\s*\}/g, "");
  s = s.replace(/\\star/g, "*").replace(/\\ast/g, "*");
  s = s.replace(/\s{2,}/g, " ").trim();
  return s;
}
function sanitizeLatexStrict(latex) {
  if (!latex) return latex;
  latex = String(latex).replace(/\s+/g, " ").trim();
  latex = latex
    .replace(/\\left(?!\s*(\(|\[|\\\{|\\langle|\\vert|\\\||\||\.))/g, "")
    .replace(/\\right(?!\s*(\)|\]|\\\}|\\rangle|\\vert|\\\||\||\.))/g, "");

  const tokens = latex.match(/\\left\b|\\right\b/g) || [];
  let bal = 0, broken = false;
  for (const t of tokens) {
    if (t === "\\left") bal++;
    else { if (bal === 0) { broken = true; break; } bal--; }
  }
  if (bal !== 0) broken = true;
  if (broken) latex = latex.replace(/\\left\s*/g, "").replace(/\\right\s*/g, "");
  return latex;
}
function fixPiecewiseFunction(latex) {
  let s = String(latex || "");
  s = s.replace(/\(\.\s+/g, "(").replace(/\s+\.\)/g, ")");
  s = s.replace(/\[\.\s+/g, "[").replace(/\s+\.\]/g, "]");
  const piecewiseMatch = s.match(/(?<!\\)\{\.\s+/);
  if (piecewiseMatch) {
    const startIdx = piecewiseMatch.index;
    const contentStart = startIdx + piecewiseMatch[0].length;
    let braceCount = 1, endIdx = contentStart, foundEnd = false;
    for (let i = contentStart; i < s.length; i++) {
      const ch = s[i];
      const prevCh = i > 0 ? s[i - 1] : "";
      if (prevCh === "\\") continue;
      if (ch === "{") braceCount++;
      else if (ch === "}") {
        braceCount--;
        if (braceCount === 0) { endIdx = i; foundEnd = true; break; }
      }
    }
    if (!foundEnd) endIdx = s.length;
    let content = s.slice(contentStart, endIdx).trim();
    content = content.replace(/\s+\.\s*$/, "");
    content = content.replace(/\s+\\\s+(?=\d)/g, " \\\\ ");
    const before = s.slice(0, startIdx);
    const after = foundEnd ? s.slice(endIdx + 1) : "";
    s = before + `\\begin{cases} ${content} \\end{cases}` + after;
  }
  return s;
}
function fixSetBracesHard(latex) {
  let s = String(latex || "");
  s = s.replace(/\\underset\s*\{([^}]*)\}\s*\{\s*l\s*i\s*m\s*\}/gi, "\\underset{$1}{\\lim}");
  s = s.replace(/\b(l)\s+(i)\s+(m)\b/gi, "lim");
  s = s.replace(/(^|[^A-Za-z\\])lim([^A-Za-z]|$)/g, "$1\\lim$2");
  s = s.replace(/\\arrow\b/g, "\\rightarrow");
  s = s.replace(/\bxarrow\b/g, "x\\rightarrow");
  s = s.replace(/\\xarrow\b/g, "\\xrightarrow");
  s = s.replace(/\\\{\s*\./g, "\\{");
  s = s.replace(/\.\s*\\\}/g, "\\}");
  s = s.replace(/\\\}\s*\./g, "\\}");
  s = s.replace(/\\mathbb\{([A-Za-z])\\\}/g, "\\mathbb{$1}");
  s = s.replace(/\\mathbb\{([A-Za-z])\}\s*\.\s*\}/g, "\\mathbb{$1}}");
  s = s.replace(/\\backslash\s*{(?!\\)/g, "\\backslash \\{");
  s = s.replace(/\\setminus\s*{(?!\\)/g, "\\setminus \\{");
  if ((s.includes("\\backslash \\{") || s.includes("\\setminus \\{")) && !s.includes("\\}")) {
    s = s.replace(/\}\s*$/g, "").trim() + "\\}";
  }
  s = s.replace(/\\\}\s*([,.;:])/g, "\\}$1");
  s = s.replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, (m, a, b) => {
    const bb = String(b).replace(/(\d)\s+(\d)/g, "$1$2");
    return `\\frac{${a}}{${bb}}`;
  });
  s = s.replace(/\s+/g, " ").trim();
  return s;
}
function restoreArrowAndCoreCommands(latex) {
  let s = String(latex || "");
  s = s.replace(/\s+/g, " ").trim();
  s = s.replace(/\b([A-Za-z])\s+arrow\b/g, "$1 \\to");
  s = s.replace(/\brightarrow\b/g, "\\rightarrow");
  s = s.replace(/\barrow\b/g, "\\rightarrow");
  s = s.replace(/(^|[^A-Za-z\\])to([^A-Za-z]|$)/g, "$1\\to$2");
  return s.replace(/\s+/g, " ").trim();
}
function normalizeLatexCommands(latex) { if (!latex) return latex; return fixSetBracesHard(String(latex)); }

function manualMathMLToLatex(mathml) {
  if (!mathml) return "";
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    textNodeName: "#text",
    preserveOrder: false,
  });

  let parsed;
  try { parsed = parser.parse(mathml); } catch { return ""; }

  function nodeToLatex(node) {
    if (!node) return "";
    if (typeof node === "string") return node;
    if (typeof node === "number") return String(node);
    if (node["#text"] !== undefined) return String(node["#text"]);
    if (Array.isArray(node)) return node.map(nodeToLatex).join("");

    let result = "";
    for (const [tag, content] of Object.entries(node)) {
      if (tag.startsWith("@_")) continue;
      const t = tag.toLowerCase();

      switch (t) {
        case "math":
        case "mrow":
        case "mstyle":
        case "mpadded":
        case "mphantom":
          result += nodeToLatex(content); break;
        case "msqrt":
          result += `\\sqrt{${nodeToLatex(content)}}`; break;
        case "mroot":
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const index = nodeToLatex(content[1]);
            result += `\\sqrt[${index}]{${base}}`;
          } else result += `\\sqrt{${nodeToLatex(content)}}`;
          break;
        case "mfrac":
          if (Array.isArray(content) && content.length >= 2) {
            result += `\\frac{${nodeToLatex(content[0])}}{${nodeToLatex(content[1])}}`;
          } else result += nodeToLatex(content);
          break;
        case "msup":
          if (Array.isArray(content) && content.length >= 2) {
            result += `${nodeToLatex(content[0])}^{${nodeToLatex(content[1])}}`;
          } else result += nodeToLatex(content);
          break;
        case "msub":
          if (Array.isArray(content) && content.length >= 2) {
            result += `${nodeToLatex(content[0])}_{${nodeToLatex(content[1])}}`;
          } else result += nodeToLatex(content);
          break;
        case "msubsup":
          if (Array.isArray(content) && content.length >= 3) {
            result += `${nodeToLatex(content[0])}_{${nodeToLatex(content[1])}}^{${nodeToLatex(content[2])}}`;
          } else result += nodeToLatex(content);
          break;
        case "mi":
        case "mn":
        case "mtext":
          result += nodeToLatex(content); break;
        case "mo": {
          const op = nodeToLatex(content);
          const opMap = {
            "√": "\\sqrt",
            "×": "\\times",
            "÷": "\\div",
            "±": "\\pm",
            "∓": "\\mp",
            "≤": "\\leq",
            "≥": "\\geq",
            "≠": "\\neq",
            "≈": "\\approx",
            "∞": "\\infty",
            "→": "\\to",
            "←": "\\leftarrow",
            "⇒": "\\Rightarrow",
            "⇐": "\\Leftarrow",
            "∈": "\\in",
            "∉": "\\notin",
            "⊂": "\\subset",
            "⊃": "\\supset",
            "∪": "\\cup",
            "∩": "\\cap",
            "∀": "\\forall",
            "∃": "\\exists",
            "∂": "\\partial",
            "∇": "\\nabla",
            "∑": "\\sum",
            "∏": "\\prod",
            "∫": "\\int",
            "α": "\\alpha",
            "β": "\\beta",
            "γ": "\\gamma",
            "δ": "\\delta",
            "ε": "\\epsilon",
            "θ": "\\theta",
            "λ": "\\lambda",
            "μ": "\\mu",
            "π": "\\pi",
            "σ": "\\sigma",
            "φ": "\\phi",
            "ω": "\\omega",
            "%": "\\%",
          };
          result += opMap[op] || op;
          break;
        }
        case "mfenced": {
          const open = node["@_open"] || "(";
          const close = node["@_close"] || ")";
          result += `\\left${open}${nodeToLatex(content)}\\right${close}`;
          break;
        }
        case "mtable":
          result += `\\begin{matrix}${nodeToLatex(content)}\\end{matrix}`; break;
        case "mtr":
          result += nodeToLatex(content) + " \\\\ "; break;
        case "mtd":
          result += nodeToLatex(content) + " & "; break;
        default:
          result += nodeToLatex(content);
      }
    }
    return result;
  }

  let latex = nodeToLatex(parsed);
  latex = latex.replace(/\s*&\s*\\\\/g, " \\\\");
  latex = latex.replace(/\s*&\s*$/g, "");
  latex = latex.replace(/\s+/g, " ").trim();
  return latex;
}

function customMathMLToLatex(mathml) {
  if (!mathml) return "";
  const pre = preprocessMathMLForSqrt(mathml);

  let latex = "";
  try { latex = MathMLToLaTeX.convert(pre) || ""; }
  catch { latex = manualMathMLToLatex(pre); }

  latex = postprocessLatexSqrt(latex);

  if (SQRT_MATHML_RE.test(mathml) && !latex.includes("\\sqrt")) {
    const manual = manualMathMLToLatex(mathml);
    if (manual.includes("\\sqrt")) return manual;
  }

  return latex.trim();
}

function mathmlToLatexSafe(mathml) {
  try { return customMathMLToLatex(mathml); }
  catch { return ""; }
}

/* ================== RELS MAP ================== */
function buildRelMaps(relsXmlText) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const rels = parser.parse(relsXmlText);
  const list = rels?.Relationships?.Relationship || [];
  const arr = Array.isArray(list) ? list : [list];

  const emb = {};
  const media = {};
  for (const r of arr) {
    const id = r?.["@_Id"];
    const target = r?.["@_Target"];
    const targetMode = r?.["@_TargetMode"];
    if (!id || !target) continue;
    if (targetMode && String(targetMode).toLowerCase() === "external") continue;

    const t = String(target).replace(/^\.?\//, "");
    const low = t.toLowerCase();
    if (low.startsWith("embeddings/") && low.endsWith(".bin")) emb[id] = "word/" + t;
    else if (low.startsWith("media/")) media[id] = "word/" + t;
  }
  return { emb, media };
}

/* ================== (KEEP) RENDER + PARSE EXAM ==================
   NOTE: To keep this answer focused on speed, phần render/exam parsing
   bạn giữ đúng code hiện có của bạn (y chang).
   -> Mình sẽ include nguyên hàm buildInlineHtml/renderParagraph/... của bạn.
   -> Vì bạn đã dán đầy đủ rồi, mình không lặp lại cho dài, nhưng bạn PHẢI
      giữ nguyên các hàm đó trong file, phía dưới phần tối ưu này.
*/
function kids(arr, tag) { return Array.isArray(arr) ? arr.filter((n) => n && typeof n === "object" && n[tag]) : []; }
function findAllRidsDeep(x, out = []) {
  const re = /^rId\d+$/;
  if (!x) return out;
  if (typeof x === "string") { const s = x.trim(); if (re.test(s)) out.push(s); return out; }
  if (Array.isArray(x)) { for (const it of x) findAllRidsDeep(it, out); return out; }
  if (typeof x === "object") { for (const v of Object.values(x)) findAllRidsDeep(v, out); return out; }
  return out;
}
function findImageEmbedRidsDeep(x, out = []) {
  if (!x) return out;
  if (Array.isArray(x)) { for (const it of x) findImageEmbedRidsDeep(it, out); return out; }
  if (typeof x === "object") {
    for (const [k, v] of Object.entries(x)) {
      if ((k === "@_r:embed" || k === "@_r:id") && typeof v === "string" && v.startsWith("rId")) out.push(v);
      findImageEmbedRidsDeep(v, out);
    }
  }
  return out;
}
function runHasOleLike(rNode) {
  try { const s = JSON.stringify(rNode); return s.includes("o:OLEObject") || s.includes("w:object") || s.includes("w:oleObject"); }
  catch { return false; }
}
function runIsUnderlined(rNode) {
  try { const s = JSON.stringify(rNode); if (!s.includes("w:u")) return false; if (s.toLowerCase().includes("none")) return false; return true; }
  catch { return false; }
}
function getTextFromPreserveWrap(tagWrap, tagName) {
  const v = tagWrap?.[tagName];
  if (!v) return "";
  if (Array.isArray(v)) return v.map((x) => x?.["#text"] || "").join("");
  if (typeof v === "object") return v?.["#text"] || "";
  return "";
}
function collectTextFromRun(rNode) {
  let s = "";
  for (const tWrap of kids(rNode, "w:t")) s += getTextFromPreserveWrap(tWrap, "w:t");
  for (const tWrap of kids(rNode, "w:instrText")) s += getTextFromPreserveWrap(tWrap, "w:instrText");
  for (const tWrap of kids(rNode, "w:delText")) s += getTextFromPreserveWrap(tWrap, "w:delText");
  if (kids(rNode, "w:tab").length) s += "\t";
  if (kids(rNode, "w:br").length) s += "\n";
  return s;
}
function escapeTextToHtml(text) {
  if (!text) return "";
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\t", "&emsp;")
    .replaceAll("\n", "<br/>");
}
function lastVisibleChar(html) {
  const t = String(html || "").replace(/<[^>]*>/g, "");
  return t.length ? t[t.length - 1] : "";
}
function appendMathWithOneSpace(html, mathSpan) {
  const prev = lastVisibleChar(html);
  if (prev && !/\s/.test(prev)) html += " ";
  html += mathSpan;
  html += " ";
  return html;
}
function renderParagraph(pNode, ctx) {
  const { latexByRid, imageByRid, debug } = ctx;
  let html = "";

  const runs = kids(pNode, "w:r");
  for (const rWrap of runs) {
    const rNode = rWrap["w:r"];
    const under = runIsUnderlined(rNode);

    if (Array.isArray(rNode)) {
      for (const child of rNode) {
        if (child["w:t"]) {
          const text = getTextFromPreserveWrap(child, "w:t");
          if (text) {
            const esc = escapeTextToHtml(text);
            html += under ? `<u>${esc}</u>` : esc;
          }
        }
        if (child["w:tab"]) html += "&emsp;";
        if (child["w:br"]) html += "<br/>";

        if (child["a:blip"] || child["pic:blipFill"] || child["w:drawing"] || child["w:pict"] || child["v:shape"]) {
          const imgRids = unique(findImageEmbedRidsDeep(child, []));
          for (const rid of imgRids) {
            const dataUri = imageByRid[rid];
            if (dataUri) { debug.imagesInjected++; html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`; }
          }
        }

        if (child["w:object"] || child["o:OLEObject"]) {
          const allRids = unique(findAllRidsDeep(child, []));
          let foundMath = false;
          for (const rid of allRids) {
            const latex = latexByRid[rid];
            if (latex) {
              debug.seenOle++;
              if (debug.sampleRids.length < 12) debug.sampleRids.push(rid);
              const mathSpan = `<span class="math">\\(${latex}\\)</span>`;
              html = appendMathWithOneSpace(html, mathSpan);
              debug.oleInjected++;
              foundMath = true;
            }
          }
          if (!foundMath) {
            const imgRids = unique(findImageEmbedRidsDeep(child, []));
            for (const rid of imgRids) {
              const dataUri = imageByRid[rid];
              if (dataUri) { debug.imagesInjected++; html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`; }
            }
          }
        }
      }
    } else {
      const runText = collectTextFromRun(rNode);
      if (runText) {
        const esc = escapeTextToHtml(runText);
        html += under ? `<u>${esc}</u>` : esc;
      }
    }

    const runImgRids = unique(findImageEmbedRidsDeep(rNode, []));
    const processedInLoop = new Set();
    if (Array.isArray(rNode)) {
      for (const child of rNode) {
        if (child["w:drawing"] || child["w:pict"] || child["v:shape"] || child["w:object"]) {
          const childRids = findImageEmbedRidsDeep(child, []);
          childRids.forEach((rid) => processedInLoop.add(rid));
        }
      }
    }
    for (const rid of runImgRids) {
      if (processedInLoop.has(rid)) continue;
      const dataUri = imageByRid[rid];
      if (dataUri) { debug.imagesInjected++; html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`; }
    }

    if (runHasOleLike(rNode)) {
      debug.seenOleRuns++;
      const rids = unique(findAllRidsDeep(rNode, []));
      const processedMathRids = new Set();
      if (Array.isArray(rNode)) {
        for (const child of rNode) {
          if (child["w:object"] || child["o:OLEObject"]) {
            const childRids = findAllRidsDeep(child, []);
            childRids.forEach((rid) => { if (latexByRid[rid]) processedMathRids.add(rid); });
          }
        }
      }
      for (const rid of rids) {
        if (processedMathRids.has(rid)) continue;
        const latex = latexByRid[rid];
        if (latex) {
          debug.seenOle++;
          if (debug.sampleRids.length < 12) debug.sampleRids.push(rid);
          const mathSpan = `<span class="math">\\(${latex}\\)</span>`;
          html = appendMathWithOneSpace(html, mathSpan);
          debug.oleInjected++;
        } else debug.ignoredRids++;
      }
    }
  }
  return html;
}
function renderTable(tblNode, ctx) {
  const rows = kids(tblNode, "w:tr");
  let html = `<table border="1" style="border-collapse:collapse;width:auto;max-width:100%;">`;
  for (const trWrap of rows) {
    const trNode = trWrap["w:tr"];
    html += "<tr>";
    const cells = kids(trNode, "w:tc");
    for (const tcWrap of cells) {
      const tcNode = tcWrap["w:tc"];
      html += `<td style="padding:6px;vertical-align:top;">`;
      const paras = kids(tcNode, "w:p");
      for (const pWrap of paras) {
        const pHtml = renderParagraph(pWrap["w:p"], ctx);
        if (pHtml) html += pHtml;
        html += "<br/>";
      }
      html += "</td>";
    }
    html += "</tr>";
  }
  html += "</table><br/>";
  return html;
}
function buildInlineHtml(documentXml, ctx) {
  const parser = new XMLParser({ ignoreAttributes: false, preserveOrder: true });
  const tree = parser.parse(documentXml);
  const doc = kids(tree, "w:document")[0]?.["w:document"];
  const body = kids(doc, "w:body")[0]?.["w:body"];
  const bodyChildren = Array.isArray(body) ? body : [];
  let html = "";
  for (const child of bodyChildren) {
    if (child["w:p"]) {
      const pHtml = renderParagraph(child["w:p"], ctx);
      if (pHtml) html += pHtml;
      html += "<br/>";
    } else if (child["w:tbl"]) {
      html += renderTable(child["w:tbl"], ctx);
    }
  }
  return html;
}

/* ================== KEEP YOUR EXISTING formatExamLayout + parseExamFromInlineHtml etc ==================
   -> Bạn copy nguyên xi từ code bạn đang có (formatExamLayout, parseExamFromInlineHtml,...)
   -> Mình không lặp lại ở đây để tránh trả lời quá dài.
*/
function splitByMath(html) {
  const out = [];
  const re = /\\\([\s\S]*?\\\)/g;
  let last = 0, m;
  while ((m = re.exec(html)) !== null) {
    if (m.index > last) out.push({ math: false, text: html.slice(last, m.index) });
    out.push({ math: true, text: m[0] });
    last = m.index + m[0].length;
  }
  if (last < html.length) out.push({ math: false, text: html.slice(last) });
  return out;
}
// ... IMPORTANT: paste ALL remaining helper functions from your current file:
// normalizeGluedChoiceMarkers, formatExamLayout, parseExamFromInlineHtml, removeUnsupportedImages, ...
// (Giữ nguyên như bạn đang có)

/* ================== ROUTES ================== */
app.get("/", (req, res) => res.type("text").send("MathType Converter API: POST /convert-docx-html, GET /health"));
app.get("/health", (req, res) => res.json({ ok: true, node: process.version, cwd: process.cwd() }));

app.post("/convert-docx-html", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) {
      return res.status(400).json({ ok: false, error: "No file uploaded. Field name must be 'file'." });
    }
    res.setHeader("Content-Type", "application/json; charset=utf-8");

    const zip = await openDocxZip(req.file.buffer);

    // fast path lookup map
    zip.__fileByPath = new Map((zip.files || []).map((f) => [f.path, f]));

    const docBuf = await readZipEntry(zip, "word/document.xml");
    const relBuf = await readZipEntry(zip, "word/_rels/document.xml.rels");
    if (!docBuf || !relBuf) {
      return res.status(400).json({ ok: false, error: "Missing word/document.xml or word/_rels/document.xml.rels" });
    }

    const { emb: embRelMap, media: mediaRelMap } = buildRelMaps(relBuf.toString("utf8"));

    const latexByRid = {};
    const mathmlByRid = {}; // only if DEBUG
    let latexOk = 0;

    // ===== OLE: fastest path with cache + persistent ruby + limiter =====
    const embEntries = Object.entries(embRelMap);

    await Promise.all(
      embEntries.map(([rid, embPath]) =>
        oleLimiter.run(async () => {
          const emb = zip.__fileByPath.get(embPath);
          if (!emb) return;

          const buf = await emb.buffer();
          const key = sha1(buf);

          const cached = oleCache.get(key);
          if (cached?.latex) {
            latexByRid[rid] = cached.latex;
            if (DEBUG && cached.mathml) mathmlByRid[rid] = cached.mathml;
            latexOk++;
            return;
          }

          // Convert OLE -> MathML via worker
          const mathml = await rubyWorker.convertOleBufferToMathML(buf);
          if (!mathml) return;

          let latex = mathmlToLatexSafe(mathml);
          if (!latex) return;

          // keep your pipeline
          latex = sanitizeLatexStrict(latex);
          latex = normalizeLatexCommands(latex);
          latex = restoreArrowAndCoreCommands(latex);
          latex = fixPiecewiseFunction(latex);
          latex = postprocessLatexSqrt(latex);
          latex = finalLatexCleanup(latex);

          oleCache.set(key, { mathml: DEBUG ? mathml : "", latex });

          latexByRid[rid] = latex;
          if (DEBUG) mathmlByRid[rid] = mathml;
          latexOk++;
        })
      )
    );

    // ===== IMAGES: cache + inkscape + limiter =====
    const imageByRid = {};
    let imagesOk = 0;
    let imagesConverted = 0;

    const mediaEntries = Object.entries(mediaRelMap);
    await Promise.all(
      mediaEntries.map(([rid, mediaPath]) =>
        imgLimiter.run(async () => {
          const mf = zip.__fileByPath.get(mediaPath);
          if (!mf) return;
          const buf = await mf.buffer();
          const ext = getExtFromPath(mediaPath);
          const key = sha1(Buffer.concat([Buffer.from(ext + ":"), buf]));

          const cached = imgCache.get(key);
          if (cached) {
            imageByRid[rid] = cached;
            imagesOk++;
            return;
          }

          if (ext === "emf" || ext === "wmf") {
            const pngBuf = convertEmfWmfToPngInkscape(buf, ext);
            if (pngBuf) {
              const dataUri = `data:image/png;base64,${pngBuf.toString("base64")}`;
              imgCache.set(key, dataUri);
              imageByRid[rid] = dataUri;
              imagesConverted++;
              imagesOk++;
            }
          } else {
            const mime = mimeFromExt(mediaPath);
            const dataUri = `data:${mime};base64,${buf.toString("base64")}`;
            imgCache.set(key, dataUri);
            imageByRid[rid] = dataUri;
            imagesOk++;
          }
        })
      )
    );

    const debug = {
      embeddings: embEntries.length,
      latexCount: Object.keys(latexByRid).length,
      latexOk,
      imagesRelCount: mediaEntries.length,
      imagesOk,
      imagesConverted,
      imagesInjected: 0,
      seenOleRuns: 0,
      seenOle: 0,
      oleInjected: 0,
      ignoredRids: 0,
      sampleRids: [],
      exam: { questions: 0, mcq: 0, tf4: 0, short: 0 },
    };

    const ctx = { latexByRid, imageByRid, debug };

    // Build HTML
    let inlineHtml = buildInlineHtml(docBuf.toString("utf8"), ctx);

    // IMPORTANT: keep your formatting pipeline here (copy from your current code)
    inlineHtml = formatExamLayout(inlineHtml);
    inlineHtml = removeUnsupportedImages(inlineHtml);

    const exam = parseExamFromInlineHtml(inlineHtml);

    if (exam) {
      debug.exam.questions = exam.questions.length;
      for (const q of exam.questions) {
        if (q.type === "mcq") debug.exam.mcq++;
        else if (q.type === "tf4") debug.exam.tf4++;
        else debug.exam.short++;
      }
    }

    return res.json({
      ok: true,
      inlineHtml,
      exam,
      debug,
      ...(DEBUG ? { mathmlByRid } : {}),
    });
  } catch (e) {
    if (DEBUG) console.error("[CONVERT_DOCX_HTML_FAIL]", e);
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
});

/* ================== START ================== */
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
