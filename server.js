import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import { execFile, execFileSync } from "child_process";
import { MathMLToLaTeX } from "mathml-to-latex";
import { XMLParser } from "fast-xml-parser";

const app = express();
app.use(cors());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/* ================= Helpers ================= */

function parseRels(relsXml) {
  const map = new Map();
  const re =
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  let m;
  while ((m = re.exec(relsXml))) map.set(m[1], m[2]);
  return map;
}

function normalizeTargetToWordPath(target) {
  let t = (target || "").replace(/^(\.\.\/)+/, "");
  if (!t.startsWith("word/")) t = `word/${t}`;
  return t;
}

function extOf(p = "") {
  return p.split(".").pop()?.toLowerCase() || "";
}

function guessMimeFromFilename(filename = "") {
  const ext = extOf(filename);
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "bmp") return "image/bmp";
  if (ext === "webp") return "image/webp";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "emf") return "image/emf";
  if (ext === "wmf") return "image/wmf";
  return "application/octet-stream";
}

function decodeXmlEntities(s = "") {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) =>
      String.fromCharCode(parseInt(h, 16))
    );
}

async function getZipEntryBuffer(zipFiles, p) {
  const f = zipFiles.find((x) => x.path === p);
  if (!f) return null;
  return await f.buffer();
}

/* ================= Inkscape Convert EMF/WMF -> PNG ================= */

function inkscapeConvertToPng(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    execFile(
      "inkscape",
      [
        inputPath,
        "--export-type=png",
        `--export-filename=${outputPath}`,
        "--export-area-drawing",
        "--export-background-opacity=0",
      ],
      { timeout: 30000 },
      (err, stdout, stderr) => {
        if (err) return reject(new Error(stderr || err.message));
        resolve(true);
      }
    );
  });
}

async function maybeConvertEmfWmfToPng(buf, filename) {
  const ext = extOf(filename);
  if (ext !== "emf" && ext !== "wmf") return null;

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "mtype-"));
  const inPath = path.join(tmpDir, `in.${ext}`);
  const outPath = path.join(tmpDir, "out.png");

  try {
    fs.writeFileSync(inPath, buf);
    await inkscapeConvertToPng(inPath, outPath);
    return fs.readFileSync(outPath);
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {}
  }
}

/* ================= MathML -> LaTeX Helpers (IMPROVED) ================= */

/**
 * Ensure MathML has namespace (some converters fail if <math> lacks xmlns)
 */
function ensureMathMLNamespace(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);

  // remove XML header if any
  s = s.replace(/<\?xml[^>]*\?>/gi, "").trim();

  // add MathML namespace if missing
  s = s.replace(
    /<math(?![^>]*\bxmlns=)/i,
    '<math xmlns="http://www.w3.org/1998/Math/MathML"'
  );

  return s;
}

/**
 * Normalize <mtable ...> -> <mtable> to avoid libraries outputting raw \\ without array/matrix
 */
function normalizeMtable(mathml) {
  if (!mathml) return mathml;
  return String(mathml).replace(/<mtable\b[^>]*>/gi, "<mtable>");
}

/**
 * Pre-process MathML to ensure sqrt elements are properly formatted
 */
function preprocessMathMLForSqrt(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);

  const moSqrt = String.raw`<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>`;

  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mrow>([\s\S]*?)<\/mrow>`, "gi"),
    "<msqrt>$1</msqrt>"
  );
  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mi>([^<]+)<\/mi>`, "gi"),
    "<msqrt><mi>$1</mi></msqrt>"
  );
  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mn>([^<]+)<\/mn>`, "gi"),
    "<msqrt><mn>$1</mn></msqrt>"
  );
  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mfenced([^>]*)>([\s\S]*?)<\/mfenced>`, "gi"),
    "<msqrt><mfenced$1>$2</mfenced></msqrt>"
  );

  return s;
}

/**
 * Post-process LaTeX to fix sqrt issues and other artifacts
 */
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

/**
 * Final LaTeX cleanup: Unicode, malformed fences, spaced functions
 */
function finalLatexCleanup(latex) {
  if (!latex) return latex;
  let s = String(latex);

  s = s.replace(/[\u200B-\u200D\uFEFF]/g, "");
  s = s.replace(/[\u00A0]/g, " ");
  s = s.replace(/[\u2000-\u200A\u202F\u205F\u3000]/g, " ");
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");

  s = s.replace(/\\left\s*\(\s*\*\s*\\right\s*\)/g, "(*)");
  s = s.replace(/\\left\s*\(\s*\\star\s*\\right\s*\)/g, "(*)");

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

  s = s.replace(/\\star/g, "*");
  s = s.replace(/\\ast/g, "*");

  s = s.replace(/\s{2,}/g, " ").trim();

  return s;
}

/**
 * Manual MathML -> LaTeX fallback (handles msqrt/mroot/mfrac etc)
 */
function manualMathMLToLatex(mathml) {
  if (!mathml) return "";

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    textNodeName: "#text",
    preserveOrder: false,
  });

  let parsed;
  try {
    parsed = parser.parse(mathml);
  } catch (e) {
    console.error("[MANUAL_PARSE_ERROR]", e?.message);
    return "";
  }

  function nodeToLatex(node) {
    if (!node) return "";
    if (typeof node === "string") return node;
    if (typeof node === "number") return String(node);

    if (node["#text"] !== undefined) {
      return String(node["#text"]);
    }

    if (Array.isArray(node)) {
      return node.map(nodeToLatex).join("");
    }

    let result = "";

    for (const [tag, content] of Object.entries(node)) {
      if (tag.startsWith("@_")) continue;
      const tagLower = tag.toLowerCase();

      switch (tagLower) {
        case "math":
        case "mrow":
        case "mstyle":
        case "mpadded":
        case "mphantom":
          result += nodeToLatex(content);
          break;

        case "msqrt":
          result += `\\sqrt{${nodeToLatex(content)}}`;
          break;

        case "mroot":
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const index = nodeToLatex(content[1]);
            result += `\\sqrt[${index}]{${base}}`;
          } else {
            result += `\\sqrt{${nodeToLatex(content)}}`;
          }
          break;

        case "mfrac":
          if (Array.isArray(content) && content.length >= 2) {
            const num = nodeToLatex(content[0]);
            const den = nodeToLatex(content[1]);
            result += `\\frac{${num}}{${den}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;

        case "msup":
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const sup = nodeToLatex(content[1]);
            result += `${base}^{${sup}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;

        case "msub":
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const sub = nodeToLatex(content[1]);
            result += `${base}_{${sub}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;

        case "msubsup":
          if (Array.isArray(content) && content.length >= 3) {
            const base = nodeToLatex(content[0]);
            const sub = nodeToLatex(content[1]);
            const sup = nodeToLatex(content[2]);
            result += `${base}_{${sub}}^{${sup}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;

        case "mi":
        case "mn":
        case "mtext":
          result += nodeToLatex(content);
          break;

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
          };
          result += opMap[op] || op;
          break;
        }

        case "mfenced": {
          const open = node["@_open"] ?? "(";
          const close = node["@_close"] ?? ")";
          result += `\\left${open}${nodeToLatex(content)}\\right${close}`;
          break;
        }

        case "mtable":
          result += `\\begin{matrix}${nodeToLatex(content)}\\end{matrix}`;
          break;

        case "mtr":
          result += nodeToLatex(content) + " \\\\ ";
          break;

        case "mtd":
          result += nodeToLatex(content) + " & ";
          break;

        default:
          result += nodeToLatex(content);
      }
    }

    return result;
  }

  let latex = nodeToLatex(parsed);

  // Clean up matrix separators
  latex = latex.replace(/\s*&\s*\\\\/g, " \\\\");
  latex = latex.replace(/\s*&\s*$/g, "");
  latex = latex.replace(/\s+/g, " ").trim();

  return latex;
}

/**
 * Try library then fallback to manual, with preprocessing/postprocessing
 */
function customMathMLToLatex(mathml) {
  if (!mathml) return "";

  let mm = ensureMathMLNamespace(mathml);
  mm = normalizeMtable(mm);
  mm = preprocessMathMLForSqrt(mm);

  let latex = "";
  try {
    latex = MathMLToLaTeX.convert(mm) || "";
  } catch (e) {
    latex = "";
  }

  if (!latex) {
    latex = manualMathMLToLatex(mm) || "";
  }

  latex = postprocessLatexSqrt(latex);
  latex = finalLatexCleanup(latex);
  return String(latex || "").trim();
}

function mathmlToLatexSafe(mml) {
  try {
    if (!mml || !mml.includes("<math")) return "";
    return customMathMLToLatex(mml);
  } catch (e) {
    console.error("[MATHML_TO_LATEX_FAIL]", e?.message || String(e));
    return "";
  }
}

/* ================= MathType OLE -> MathML -> LaTeX ================= */

/**
 * scan MathML embedded directly in OLE (nhanh)
 */
function extractMathMLFromOleScan(buf) {
  const utf8 = buf.toString("utf8");
  let i = utf8.indexOf("<math");
  if (i !== -1) {
    let j = utf8.indexOf("</math>", i);
    if (j !== -1) return utf8.slice(i, j + 7);
  }

  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    let j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }

  return null;
}

/**
 * fallback: call ruby mt2mml.rb ole.bin -> MathML
 */
function rubyOleToMathML(oleBuf) {
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "ole-"));
    const inPath = path.join(tmpDir, "oleObject.bin");
    fs.writeFileSync(inPath, oleBuf);

    execFile(
      "ruby",
      ["mt2mml.rb", inPath],
      { timeout: 30000, maxBuffer: 20 * 1024 * 1024 },
      (err, stdout, stderr) => {
        try {
          fs.rmSync(tmpDir, { recursive: true, force: true });
        } catch {}
        if (err) return reject(new Error(stderr || err.message));
        resolve(String(stdout || "").trim());
      }
    );
  });
}

/**
 * MathType FIRST:
 * - token [!m:$mathtype_x$]
 * - produce:
 *   latexMap[key] = "..."
 *   (optional) images["fallback_key"] = png dataURL if latex fails
 *
 * NOTE: giữ nguyên thuật toán, chỉ dùng mathmlToLatexSafe (mạnh hơn) khi chuyển MathML -> LaTeX
 */
async function tokenizeMathTypeOleFirst(docXml, rels, zipFiles, images) {
  let idx = 0;
  const found = {}; // key -> { oleTarget, previewRid }

  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;

  docXml = docXml.replace(OBJECT_RE, (block) => {
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid);
    if (!oleTarget) return block;

    const vmlRid = block.match(/<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/);
    const blipRid = block.match(/<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/);
    const previewRid = vmlRid?.[1] || blipRid?.[1] || null;

    const key = `mathtype_${++idx}`;
    found[key] = { oleTarget, previewRid };
    return `[!m:$${key}$]`;
  });

  const latexMap = {};

  await Promise.all(
    Object.entries(found).map(async ([key, info]) => {
      const oleFull = normalizeTargetToWordPath(info.oleTarget);
      const oleBuf = await getZipEntryBuffer(zipFiles, oleFull);

      // 1) try scan MathML inside OLE
      let mml = "";
      if (oleBuf) mml = extractMathMLFromOleScan(oleBuf) || "";

      // 2) fallback ruby convert (MTEF inside OLE)
      if (!mml && oleBuf) {
        try {
          mml = await rubyOleToMathML(oleBuf);
        } catch {
          mml = "";
        }
      }

      // 3) MathML -> LaTeX (IMPROVED)
      const latex = mml ? mathmlToLatexSafe(mml) : "";
      if (latex) {
        latexMap[key] = latex;
        return;
      }

      // 4) If no latex, fallback to preview image (convert emf/wmf->png)
      if (info.previewRid) {
        const t = rels.get(info.previewRid);
        if (t) {
          const imgFull = normalizeTargetToWordPath(t);
          const imgBuf = await getZipEntryBuffer(zipFiles, imgFull);
          if (imgBuf) {
            const mime = guessMimeFromFilename(imgFull);
            if (mime === "image/emf" || mime === "image/wmf") {
              try {
                const pngBuf = await maybeConvertEmfWmfToPng(imgBuf, imgFull);
                if (pngBuf) {
                  images[`fallback_${key}`] =
                    `data:image/png;base64,${pngBuf.toString("base64")}`;
                  latexMap[key] = "";
                  return;
                }
              } catch {}
            }
            images[`fallback_${key}`] =
              `data:${mime};base64,${imgBuf.toString("base64")}`;
          }
        }
      }

      latexMap[key] = "";
    })
  );

  return { outXml: docXml, latexMap };
}

/**
 * Tokenize normal images AFTER MathType (and convert EMF/WMF -> PNG if possible)
 */
async function tokenizeImagesAfter(docXml, rels, zipFiles) {
  let idx = 0;
  const imgMap = {};
  const jobs = [];

  const schedule = (rid, key) => {
    const target = rels.get(rid);
    if (!target) return;
    const full = normalizeTargetToWordPath(target);

    jobs.push(
      (async () => {
        const buf = await getZipEntryBuffer(zipFiles, full);
        if (!buf) return;

        const mime = guessMimeFromFilename(full);
        if (mime === "image/emf" || mime === "image/wmf") {
          try {
            const pngBuf = await maybeConvertEmfWmfToPng(buf, full);
            if (pngBuf) {
              imgMap[key] = `data:image/png;base64,${pngBuf.toString("base64")}`;
              return;
            }
          } catch {}
        }
        imgMap[key] = `data:${mime};base64,${buf.toString("base64")}`;
      })()
    );
  };

  docXml = docXml.replace(
    /<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/g,
    (m, rid) => {
      const key = `img_${++idx}`;
      schedule(rid, key);
      return `[!img:$${key}$]`;
    }
  );

  docXml = docXml.replace(
    /<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/g,
    (m, rid) => {
      const key = `img_${++idx}`;
      schedule(rid, key);
      return `[!img:$${key}$]`;
    }
  );

  await Promise.all(jobs);
  return { outXml: docXml, imgMap };
}

/* ================= Text & Questions ================= */
/**
 * ✅ HOÀN THIỆN:
 * - GIỮ token [!m:$...$]/[!img:$...$] để HTML render công thức + ảnh
 * - GIỮ underline bằng <u>...</u>
 * - Không đổi thuật toán MathType/LaTeX phía trên
 */
function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // 1) Protect tokens BEFORE stripping tags (both $key$ and $$key$$)
  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

  // 2) Convert each run <w:r> while preserving underline + token text
  x = x.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
    const hasU =
      /<w:u\b[^>]*\/>/.test(run) &&
      !/<w:u\b[^>]*w:val="none"[^>]*\/>/.test(run);

    // remove run properties but keep content (tokens are plain text)
    let inner = run.replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g, "");

    // w:t -> raw text
    inner = inner.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (_, t) => t ?? "");

    // optional instrText
    inner = inner.replace(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g,
      (_, t) => t ?? ""
    );

    // remove remaining tags inside run
    inner = inner.replace(/<[^>]+>/g, "");

    if (!inner) return "";
    return hasU ? `<u>${inner}</u>` : inner;
  });

  // 3) Remove remaining tags outside runs, but keep <u>
  x = x.replace(/<(?!\/?u\b)[^>]+>/g, "");

  // 4) Restore tokens in a stable form (use $$ ... $$)
  x = x
    .replace(/___MATH_TOKEN___(.*?)___END___/g, "[!m:$$$1$$]")
    .replace(/___IMG_TOKEN___(.*?)___END___/g, "[!img:$$$1$$]");

  x = decodeXmlEntities(x)
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return x;
}

function parseQuestions(text) {
  const blocks = text.split(/(?=Câu\s+\d+\.)/);
  const questions = [];

  for (const block of blocks) {
    if (!block.startsWith("Câu")) continue;

    const q = {
      type: "multiple_choice",
      content: "",
      choices: [],
      correct: null,
      solution: "",
    };

    const [main, solution] = block.split(/Lời giải/i);
    q.solution = solution ? solution.trim() : "";

    // ✅ giữ nguyên thuật toán parse đáp án của bạn
    const choiceRe =
      /(\*?)([A-D])\.\s([\s\S]*?)(?=\n\*?[A-D]\.\s|\nLời giải|$)/g;

    let m;
    while ((m = choiceRe.exec(main))) {
      const starred = m[1] === "*";
      const label = m[2];
      const content = (m[3] || "").trim();
      if (starred) q.correct = label;
      q.choices.push({ label, text: content });
    }

    const splitAtA = main.split(/\n\*?A\.\s/);
    q.content = splitAtA[0].trim();
    questions.push(q);
  }

  return questions;
}

/* ================= Section extraction (non-intrusive) ================= */

/**
 * Extract simple "PHẦN ..." or "PHẦN X." style headers from document.xml
 * This is non-invasive: chỉ đọc documentXml để trả về sections array, không thay đổi tokenization/parsing.
 */
function extractSectionsFromDocXml(docXml) {
  if (!docXml) return [];
  const re = /<w:t[^>]*>([^<]*?(?:PHẦN|Phần|PHAN|Phần)\s*\d[^<]*)<\/w:t>/gi;
  const out = [];
  let m;
  while ((m = re.exec(docXml)) !== null) {
    const txt = decodeXmlEntities(m[1].trim());
    out.push({ title: txt });
  }
  // unique
  return [...new Map(out.map(s => [s.title, s])).values()];
}

/* ================= API ================= */

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) throw new Error("No file uploaded");

    const zip = await unzipper.Open.buffer(req.file.buffer);

    const docEntry = zip.files.find((f) => f.path === "word/document.xml");
    const relEntry = zip.files.find(
      (f) => f.path === "word/_rels/document.xml.rels"
    );
    if (!docEntry || !relEntry)
      throw new Error("Missing document.xml or document.xml.rels");

    let docXml = (await docEntry.buffer()).toString("utf8");
    const relsXml = (await relEntry.buffer()).toString("utf8");
    const rels = parseRels(relsXml);

    // 1) MathType -> LaTeX (and fallback images) ✅ giữ nguyên thuật toán
    const images = {};
    const mt = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mt.outXml;
    const latexMap = mt.latexMap;

    // 2) normal images ✅ giữ nguyên
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // 3) text ✅ giờ giữ token + underline
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) parse questions ✅ giữ nguyên
    const questions = parseQuestions(text);

    // 5) extract sections (không thay đổi thuật toán chính)
    const sections = extractSectionsFromDocXml((await docEntry.buffer()).toString("utf8"));

    res.json({
      ok: true,
      total: questions.length,
      questions,
      latex: latexMap,
      images,
      rawText: text,
      sections, // bổ sung thông tin tiêu đề phần (nếu có)
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

app.get("/debug-inkscape", (_, res) => {
  try {
    const v = execFileSync("inkscape", ["--version"]).toString();
    res.type("text/plain").send(v);
  } catch {
    res.status(500).type("text/plain").send("NO INKSCAPE");
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
