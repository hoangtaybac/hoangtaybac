// server da bo sung phan hoan chinh theo goc.js
// Merged preserve-order renderParagraph into original server flow
// Usage: node "server da bo sung phan hoan chinh theo goc.js"
// Requires: ruby + mt2mml.rb (or mt2mml_v2.rb), libreoffice/imagemagick/inkscape optional
// npm i express multer unzipper cors fast-xml-parser mathml-to-latex p-limit

import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import { execFileSync, execFile, execSync } from "child_process";
import { XMLParser } from "fast-xml-parser";
import { MathMLToLaTeX } from "mathml-to-latex";
import crypto from "crypto";
import pLimit from "p-limit";

const app = express();
app.use(cors());
app.use(express.json({ limit: "25mb" }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/* ================= Utilities ================= */
function safeUnlink(p) {
  try { fs.unlinkSync(p); } catch {}
}
function safeRmdir(dir) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch {}
}
function uniqueTmpPath(baseName = "oleObject.bin") {
  const safe = path.basename(baseName).replace(/[^\w.\-]/g, "_");
  return path.join(
    os.tmpdir(),
    `${Date.now()}_${Math.random().toString(16).slice(2)}_${safe}`
  );
}
function getExtFromPath(p) {
  return (p.split(".").pop() || "").toLowerCase();
}
function mimeFromExt(p) {
  const ext = getExtFromPath(p);
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "emf") return "image/emf";
  if (ext === "wmf") return "image/wmf";
  return "application/octet-stream";
}

/* ================== ZIP helpers ================== */
async function openDocxZip(docxBuffer) {
  return unzipper.Open.buffer(docxBuffer);
}
async function readZipEntry(zip, p) {
  const f = (zip.files || []).find((x) => x.path === p);
  if (!f) return null;
  return await f.buffer();
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

/* ================== EMF/WMF -> PNG CONVERSION ================== */
function convertEmfWmfToPng(buffer, ext) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "img-convert-"));
  const inFile = path.join(tmpDir, `input.${ext}`);

  try {
    fs.writeFileSync(inFile, buffer);

    try {
      execSync(
        `soffice --headless --convert-to png "${inFile}" --outdir "${tmpDir}"`,
        { stdio: "ignore", timeout: 30000 }
      );

      const pngFile = fs.readdirSync(tmpDir).find(f => f.endsWith(".png"));
      if (pngFile) {
        return fs.readFileSync(path.join(tmpDir, pngFile));
      }
    } catch (loErr) {
      // ignore
    }

    try {
      const outFile = path.join(tmpDir, "output.png");
      execSync(`convert "${inFile}" "${outFile}"`, { stdio: "ignore", timeout: 30000 });
      if (fs.existsSync(outFile)) {
        return fs.readFileSync(outFile);
      }
    } catch (imErr) {
      // ignore
    }

    return null;
  } catch (e) {
    return null;
  } finally {
    safeRmdir(tmpDir);
  }
}

/* ================== RUBY OLE(.bin) -> MATHML (with cache & concurrency) ================== */
const oleCache = new Map();
const rubyLimit = pLimit(3);

function rubyConvertOleBinToMathML(oleBinBuffer, filenameForTmp) {
  const tmpPath = uniqueTmpPath(path.basename(filenameForTmp || "oleObject.bin"));
  fs.writeFileSync(tmpPath, oleBinBuffer);

  try {
    const v2Script = path.join(process.cwd(), "mt2mml_v2.rb");
    const v1Script = path.join(process.cwd(), "mt2mml.rb");
    const scriptToUse = fs.existsSync(v2Script) ? v2Script : v1Script;

    const out = execFileSync(
      "ruby",
      [scriptToUse, tmpPath],
      { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], timeout: 30000 }
    );

    let mathml = "";
    try {
      const parsed = JSON.parse(out);
      mathml = parsed.mathml || "";
    } catch {
      mathml = (out || "").trim();
    }

    if (!mathml || !mathml.startsWith("<")) return "";
    return mathml;
  } catch (e) {
    return "";
  } finally {
    safeUnlink(tmpPath);
  }
}

async function rubyConvertWithCache(buf, filename) {
  const key = crypto.createHash("sha1").update(buf).digest("hex");
  if (oleCache.has(key)) {
    const v = oleCache.get(key);
    return typeof v === "string" ? v : await v;
  }
  const p = rubyLimit(() => rubyConvertOleBinToMathML(buf, filename));
  oleCache.set(key, p);
  const res = await p;
  oleCache.set(key, res || "");
  return res;
}

/* ================== MATHML -> LATEX (safe) ================== */
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
  s = s.replace(/\\surd\b/g, '\\sqrt{}');
  s = s.replace(/√\s*\{([^}]+)\}/g, '\\sqrt{$1}');
  s = s.replace(/√\s*\(([^)]+)\)/g, '\\sqrt{$1}');
  s = s.replace(/√\s*(\d+)/g, '\\sqrt{$1}');
  s = s.replace(/√\s*([a-zA-Z])/g, '\\sqrt{$1}');
  s = s.replace(/\\sqrt\s+(\d+)(?![}\d])/g, '\\sqrt{$1}');
  s = s.replace(/\\sqrt\s+([a-zA-Z])(?![}\w])/g, '\\sqrt{$1}');
  s = s.replace(/\\sqrt\s*\{\s*\}/g, '\\sqrt{\\phantom{x}}');
  s = s.replace(/\\root\s*\{([^}]+)\}\s*\\of\s*\{([^}]+)\}/g, '\\sqrt[$1]{$2}');
  s = s.replace(/\\sqrt\s*\[\s*(\d+)\s*\]\s*\{/g, '\\sqrt[$1]{');
  return s;
}

function manualMathMLToLatex(mathml) {
  // lightweight manual fallback focusing on sqrt/mroot/mfrac
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
    if (node["#text"] !== undefined) return String(node["#text"]);
    if (Array.isArray(node)) return node.map(nodeToLatex).join("");
    let res = "";
    for (const [k, v] of Object.entries(node)) {
      const tag = k.toLowerCase();
      switch (tag) {
        case "msqrt": res += `\\sqrt{${nodeToLatex(v)}}`; break;
        case "mroot":
          if (Array.isArray(v) && v.length >= 2) {
            res += `\\sqrt[${nodeToLatex(v[1])}]{${nodeToLatex(v[0])}}`;
          } else res += `\\sqrt{${nodeToLatex(v)}}`;
          break;
        case "mfrac":
          if (Array.isArray(v) && v.length >= 2) res += `\\frac{${nodeToLatex(v[0])}}{${nodeToLatex(v[1])}}`;
          else res += nodeToLatex(v);
          break;
        case "mi": case "mn": case "mtext": res += nodeToLatex(v); break;
        case "mo": res += (nodeToLatex(v) || ""); break;
        case "mrow": res += nodeToLatex(v); break;
        default: res += nodeToLatex(v);
      }
    }
    return res;
  }
  let out = nodeToLatex(parsed);
  return out.replace(/\s+/g, " ").trim();
}

function mathmlToLatexSafe(mathml) {
  if (!mathml || !mathml.includes("<math")) return "";
  try {
    const pre = preprocessMathMLForSqrt(mathml);
    let latex = (MathMLToLaTeX.convert(pre) || "").trim();
    if (!latex) latex = manualMathMLToLatex(pre);
    latex = postprocessLatexSqrt(latex);
    return latex.trim();
  } catch {
    return manualMathMLToLatex(mathml);
  }
}

/* ================== PRESERVE-ORDER HELPERS (renderParagraph, renderTable, buildInlineHtml) ================== */
function kids(arr, tag) {
  return Array.isArray(arr) ? arr.filter((n) => n && typeof n === "object" && n[tag]) : [];
}
function findAllRidsDeep(x, out = []) {
  const re = /^rId\d+$/;
  if (!x) return out;
  if (typeof x === "string") {
    const s = x.trim();
    if (re.test(s)) out.push(s);
    return out;
  }
  if (Array.isArray(x)) {
    for (const it of x) findAllRidsDeep(it, out);
    return out;
  }
  if (typeof x === "object") {
    for (const v of Object.values(x)) findAllRidsDeep(v, out);
    return out;
  }
  return out;
}
function findImageEmbedRidsDeep(x, out = []) {
  if (!x) return out;
  if (Array.isArray(x)) {
    for (const it of x) findImageEmbedRidsDeep(it, out);
    return out;
  }
  if (typeof x === "object") {
    for (const [k, v] of Object.entries(x)) {
      if ((k === "@_r:embed" || k === "@_r:id") && typeof v === "string" && v.startsWith("rId")) out.push(v);
      findImageEmbedRidsDeep(v, out);
    }
  }
  return out;
}
function runHasOleLike(rNode) {
  try {
    const s = JSON.stringify(rNode);
    return s.includes("o:OLEObject") || s.includes("w:object") || s.includes("w:oleObject");
  } catch { return false; }
}
function runIsUnderlined(rNode) {
  try {
    const s = JSON.stringify(rNode);
    if (!s.includes("w:u")) return false;
    if (s.toLowerCase().includes("none")) return false;
    return true;
  } catch { return false; }
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

        if (child["a:blip"] || child["pic:blipFill"] || child["w:drawing"]) {
          const imgRids = Array.from(new Set(findImageEmbedRidsDeep(child, [])));
          for (const rid of imgRids) {
            const dataUri = imageByRid[rid];
            if (dataUri) {
              debug.imagesInjected++;
              html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
            }
          }
        }

        if (child["w:pict"] || child["v:shape"]) {
          const imgRids = Array.from(new Set(findImageEmbedRidsDeep(child, [])));
          for (const rid of imgRids) {
            const dataUri = imageByRid[rid];
            if (dataUri) {
              debug.imagesInjected++;
              html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
            }
          }
        }

        if (child["w:object"] || child["o:OLEObject"]) {
          const allRids = Array.from(new Set(findAllRidsDeep(child, [])));

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
            const imgRids = Array.from(new Set(findImageEmbedRidsDeep(child, [])));
            for (const rid of imgRids) {
              const dataUri = imageByRid[rid];
              if (dataUri) {
                debug.imagesInjected++;
                html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
              }
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

    const runImgRids = Array.from(new Set(findImageEmbedRidsDeep(rNode, [])));
    const processedInLoop = new Set();

    if (Array.isArray(rNode)) {
      for (const child of rNode) {
        if (child["w:drawing"] || child["w:pict"] || child["v:shape"] || child["w:object"]) {
          const childRids = findImageEmbedRidsDeep(child, []);
          childRids.forEach(rid => processedInLoop.add(rid));
        }
      }
    }

    for (const rid of runImgRids) {
      if (processedInLoop.has(rid)) continue;
      const dataUri = imageByRid[rid];
      if (dataUri) {
        debug.imagesInjected++;
        html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
      }
    }

    if (runHasOleLike(rNode)) {
      debug.seenOleRuns++;
      const rids = Array.from(new Set(findAllRidsDeep(rNode, [])));

      const processedMathRids = new Set();
      if (Array.isArray(rNode)) {
        for (const child of rNode) {
          if (child["w:object"] || child["o:OLEObject"]) {
            const childRids = findAllRidsDeep(child, []);
            childRids.forEach(rid => {
              if (latexByRid[rid]) processedMathRids.add(rid);
            });
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
        } else {
          debug.ignoredRids++;
        }
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

/* ================== FORMAT LAYOUT + PARSE EXAM (from inlineHtml) ================== */
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

function normalizeGluedChoiceMarkers(s) {
  s = String(s || "");
  s = s.replace(/([^<\s>])([ABCD])\./g, "$1 $2.");
  s = s.replace(/([^<\s>])([a-d])\)/gi, "$1 $2)");
  s = s.replace(/([^<\s>])(<u[^>]*>\s*[ABCD]\s*<\/u>\s*\.)/gi, "$1 $2");
  s = s.replace(/([^<\s>])(<u[^>]*>\s*[a-d]\s*<\/u>\s*\))/gi, "$1 $2");
  return s;
}

function formatAbcdOutsideHeaders(text) {
  const headerRegex = /(<div class="section-header">[\s\S]*?<\/div>)/g;
  const segments = text.split(headerRegex);

  return segments.map(seg => {
    if (seg.startsWith('<div class="section-header">')) {
      return seg;
    }
    let s = seg;
    s = s
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*([a-d])\)/gi, "$1&emsp;$2)")
      .replace(/([^<\n])\s*([a-d])\)/gi, "$1<br/>&emsp;$2)");
    s = s
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*(<u[^>]*>\s*[a-d]\s*\)\s*<\/u>)/gi, "$1&emsp;$2")
      .replace(/([^<\n])\s*(<u[^>]*>\s*[a-d]\s*\)\s*<\/u>)/gi, "$1<br/>&emsp;$2");
    s = s
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*(<u[^>]*>\s*[a-d]\s*<\/u>\s*\))/gi, "$1&emsp;$2")
      .replace(/([^<\n])\s*(<u[^>]*>\s*[a-d]\s*<\/u>\s*\))/gi, "$1<br/>&emsp;$2");
    return s;
  }).join('');
}

function formatExamLayout(html) {
  let result = html;
  result = result.replace(/\s+/g, " ");
  result = result.replace(/PHẦN(\d)/gi, "PHẦN $1");
  result = result.replace(
    /(^|<br\/>)\s*(PHẦN\s+\d+\.(?:(?!<br\/>\s*Câu\s+\d).)*)/g,
    '$1<br/><div class="section-header"><strong>$2</strong></div>'
  );
  const parts = splitByMath(result);

  for (const p of parts) {
    if (p.math) continue;
    p.text = normalizeGluedChoiceMarkers(p.text);
    p.text = p.text
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*([ABCD])\./g, "$1&emsp;$2.")
      .replace(/([^<\n])\s*([ABCD])\./g, "$1<br/>&emsp;$2.");

    p.text = p.text
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*(<u[^>]*>\s*[ABCD]\s*<\/u>\s*\.)/gi, "$1&emsp;$2")
      .replace(/([^<\n])\s*(<u[^>]*>\s*[ABCD]\s*<\/u>\s*\.)/gi, "$1<br/>&emsp;$2");

    p.text = formatAbcdOutsideHeaders(p.text);

    p.text = p.text.replace(/(Câu)\s*(\d+)\s*\./g, "$1 $2.");
    p.text = p.text.replace(/(<br\/>\s*){3,}/g, "<br/><br/>");
  }

  return parts.map((x) => x.text).join("");
}

function stripAllTagsToPlain(html) {
  return String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/&emsp;/g, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractUnderlinedKeys(blockHtml) {
  const keys = { mcq: null, tf: [] };
  const s = String(blockHtml || "");
  let m =
    s.match(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./i) ||
    s.match(/<u[^>]*>\s*([A-D])\.\s*<\/u>/i);
  if (m) keys.mcq = m[1].toUpperCase();
  let mm;
  const reTF1 = /<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi;
  while ((mm = reTF1.exec(s)) !== null) keys.tf.push(mm[1].toLowerCase());
  const reTF2 = /<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi;
  while ((mm = reTF2.exec(s)) !== null) keys.tf.push(mm[1].toLowerCase());
  keys.tf = [...new Set(keys.tf)];
  return keys;
}

function normalizeUnderlinedMarkersForSplit(html) {
  let s = String(html || "");
  s = s.replace(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./gi, "$1.");
  s = s.replace(/<u[^>]*>\s*([A-D])\.\s*<\/u>/gi, "$1.");
  s = s.replace(/<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi, "$1)");
  s = s.replace(/<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi, "$1)");
  return s;
}

function removeUnsupportedImages(html) {
  let s = String(html || "");
  s = s.replace(/<img[^>]*src\s*=\s*["']\s*["'][^>]*>/gi, "");
  s = s.replace(/<img(?![^>]*src\s*=)[^>]*>/gi, "");
  s = s.replace(/<img[^>]*data:application\/octet-stream[^>]*>/gi, "");
  return s;
}

function splitChoicesHtmlABCD(blockHtml) {
  let s = normalizeUnderlinedMarkersForSplit(blockHtml);
  s = s.replace(/&emsp;/g, " ");
  s = normalizeGluedChoiceMarkers(s);
  s = s.replace(/<br\/>/g, " <br/>");

  const re = /(^|[\s>.:;,<\)\]\}！？\?])([ABCD])\./g;

  const hits = [];
  let m;
  while ((m = re.exec(s)) !== null) hits.push({ idx: m.index + m[1].length, key: m[2] });
  if (hits.length < 2) return null;

  const lastStart = hits[hits.length - 1].idx;
  const solIdx = findSolutionMarkerIndex(s, lastStart);
  const endAll = solIdx >= 0 ? solIdx : s.length;

  const out = {
    _stem: s.slice(0, hits[0].idx).trim(),
    _tail: solIdx >= 0 ? s.slice(solIdx).trim() : "",
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : endAll;
    let seg = s.slice(start, end).trim();
    seg = seg.replace(/^([ABCD])\.\s*/i, "");
    out[key] = removeUnsupportedImages(seg.trim());
  }
  return out;
}

function splitStatementsHtmlabcd(blockHtml) {
  let s = normalizeUnderlinedMarkersForSplit(blockHtml);
  s = s.replace(/&emsp;/g, " ");
  s = normalizeGluedChoiceMarkers(s);
  s = s.replace(/<br\/>/g, " <br/>");

  const earlysolIdx = findSolutionMarkerIndex(s, 0);
  let workingHtml = s;
  let tailHtml = "";

  if (earlysolIdx >= 0) {
    workingHtml = s.slice(0, earlysolIdx);
    tailHtml = s.slice(earlysolIdx).trim();
  }

  const re = /(^|[\s>.:;,<\)\]\}！？\?])([a-d])\)/gi;

  const hits = [];
  let m;
  while ((m = re.exec(workingHtml)) !== null) {
    hits.push({ idx: m.index + m[1].length, key: m[2].toLowerCase() });
  }
  if (hits.length < 2) return null;

  const out = {
    _stem: workingHtml.slice(0, hits[0].idx).trim(),
    _tail: tailHtml,
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : workingHtml.length;
    let seg = workingHtml.slice(start, end).trim();
    seg = seg.replace(/^([a-d])\)\s*/i, "");
    out[key] = removeUnsupportedImages(seg.trim());
  }
  return out;
}

function findSolutionMarkerIndex(html, fromIndex = 0) {
  const s = String(html || "");
  const re = /(Lời(?:\s*<[^>]*>)*\s*giải|Giải(?:\s*<[^>]*>)*\s*chi\s*tiết|Hướng(?:\s*<[^>]*>)*\s*dẫn(?:\s*<[^>]*>)*\s*giải)/i;
  const sub = s.slice(fromIndex);
  const m = re.exec(sub);
  if (!m) return -1;
  return fromIndex + m.index;
}

function splitSolutionSections(tailHtml) {
  let s = String(tailHtml || "").trim();
  if (!s) return { solutionHtml: "", detailHtml: "" };
  const reCT = /(Giải(?:\s*<[^>]*>)*\s*chi\s*tiết)/i;
  const matchCT = reCT.exec(s);
  if (matchCT) {
    const idxCT = matchCT.index;
    return {
      solutionHtml: s.slice(0, idxCT).trim(),
      detailHtml: s.slice(idxCT).trim(),
    };
  }
  return { solutionHtml: s, detailHtml: "" };
}

function cleanStem(html) {
  if (!html) return html;
  return String(html).replace(/^Câu\s+\d+\.?\s*/i, '').trim();
}

function parseExamFromInlineHtml(inlineHtml) {
  const re = /(^|<br\/>\s*)\s*(?:<[^>]*>\s*)*Câu\s+(\d+)\./gi;
  const hits = [];
  let m;
  while ((m = re.exec(inlineHtml)) !== null) {
    const startAt = m.index + (m[1] ? m[1].length : 0);
    hits.push({ qno: Number(m[2]), pos: startAt });
  }
  if (!hits.length) return null;

  const sectionRe = /<div class="section-header"><strong>([\s\S]*?)<\/strong><\/div>/gi;
  const sections = [];
  let sectionMatch;
  while ((sectionMatch = sectionRe.exec(inlineHtml)) !== null) {
    sections.push({
      pos: sectionMatch.index,
      html: sectionMatch[0],
      title: sectionMatch[1].trim()
    });
  }

  const rawBlocks = [];
  for (let i = 0; i < hits.length; i++) {
    const start = hits[i].pos;
    let end = i + 1 < hits.length ? hits[i + 1].pos : inlineHtml.length;
    for (const sec of sections) {
      if (sec.pos > start && sec.pos < end) {
        end = sec.pos;
        break;
      }
    }
    rawBlocks.push({ qno: hits[i].qno, pos: hits[i].pos, html: inlineHtml.slice(start, end) });
  }

  const blocks = [];
  for (const b of rawBlocks) {
    const last = blocks[blocks.length - 1];
    if (last && last.qno === b.qno) {
      last.html += "<br/>" + b.html;
    } else {
      blocks.push({ ...b });
    }
  }

  const exam = { version: 8, questions: [], sections };

  function findSectionForQuestion(qPos) {
    let currentSection = null;
    for (const sec of sections) {
      if (sec.pos < qPos) {
        currentSection = sec;
      } else {
        break;
      }
    }
    return currentSection;
  }

  for (const b of blocks) {
    const under = extractUnderlinedKeys(b.html);
    const plain = stripAllTagsToPlain(b.html);
    const section = findSectionForQuestion(b.pos);

    const isMCQ = /\b[ABCD]\./.test(plain) && (plain.match(/\b[ABCD]\./g) || []).length >= 2;
    const isTF4 = !isMCQ && (plain.match(/\b[a-d]\)/gi) || []).length >= 2;

    if (isMCQ) {
      const parts = splitChoicesHtmlABCD(b.html);
      const sol = splitSolutionSections(parts?._tail || "");
      exam.questions.push({
        no: b.qno,
        type: "mcq",
        stemHtml: cleanStem(parts?._stem || b.html),
        choicesHtml: { A: parts?.A || "", B: parts?.B || "", C: parts?.C || "", D: parts?.D || "" },
        answer: under.mcq,
        solutionHtml: sol.solutionHtml,
        detailHtml: sol.detailHtml,
        _plain: plain,
        section: section ? { title: section.title, html: section.html } : null
      });
      continue;
    }

    if (isTF4) {
      const parts = splitStatementsHtmlabcd(b.html);
      const sol = splitSolutionSections(parts?._tail || "");
      const ans = { a: null, b: null, c: null, d: null };
      for (const k of ["a", "b", "c", "d"]) {
        if (under.tf.includes(k)) ans[k] = true;
      }
      exam.questions.push({
        no: b.qno,
        type: "tf4",
        stemHtml: cleanStem(parts?._stem || b.html),
        statements: { a: parts?.a || "", b: parts?.b || "", c: parts?.c || "", d: parts?.d || "" },
        answer: ans,
        solutionHtml: sol.solutionHtml,
        detailHtml: sol.detailHtml,
        _plain: plain,
        section: section ? { title: section.title, html: section.html } : null
      });
      continue;
    }

    const solIdx = findSolutionMarkerIndex(b.html, 0);
    const stemPart = solIdx >= 0 ? b.html.slice(0, solIdx).trim() : b.html;
    const tailPart = solIdx >= 0 ? b.html.slice(solIdx).trim() : "";
    const sol = splitSolutionSections(tailPart);

    exam.questions.push({
      no: b.qno,
      type: "short",
      stemHtml: cleanStem(stemPart),
      boxes: 4,
      solutionHtml: sol.solutionHtml || tailPart,
      detailHtml: sol.detailHtml || "",
      _plain: plain,
      section: section ? { title: section.title, html: section.html } : null
    });
  }

  return exam;
}

/* ================== attachSectionOrder + buildOrderedBlocks ================== */
function attachSectionOrderToQuestions(exam, sections) {
  if (!exam?.questions?.length || !Array.isArray(sections)) return;
  for (const q of exam.questions) {
    q.sectionOrder = null;
    q.sectionTitle = null;
  }
  for (let i = 0; i < sections.length; i++) {
    const sec = sections[i];
    if (!sec) continue;
    // determine questionIndexStart/End by scanning exam.questions positions (we don't have raw positions here)
    // We'll derive by matching section.title presence in question.section if present
    for (let qi = 0; qi < exam.questions.length; qi++) {
      const q = exam.questions[qi];
      if (q.section && q.section.title === sec.title) {
        // back-fill contiguous questions belonging to same section (approx)
        q.sectionOrder = i + 1;
        q.sectionTitle = sec.title;
      }
    }
  }
  // If some questions still null, set to 1
  for (const q of exam.questions) {
    if (!q.sectionOrder) {
      q.sectionOrder = 1;
      if (!q.sectionTitle) q.sectionTitle = "PHẦN 1";
    }
  }
}

function buildOrderedBlocksFromExam(exam) {
  const blocks = [];
  let lastSec = null;
  for (const q of exam?.questions || []) {
    const sec = q.sectionOrder || null;
    if (sec && sec !== lastSec) {
      blocks.push({
        type: "section",
        order: sec,
        title: q.sectionTitle || `PHẦN ${sec}`,
      });
      lastSec = sec;
    }
    blocks.push({ type: "question", data: q });
  }
  return blocks;
}

/* ================== ROUTE: /upload (main) ================== */
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) throw new Error("No file uploaded");
    const zip = await openDocxZip(req.file.buffer);

    const docBuf = await readZipEntry(zip, "word/document.xml");
    const relBuf = await readZipEntry(zip, "word/_rels/document.xml.rels");
    if (!docBuf || !relBuf) throw new Error("Missing document.xml or document.xml.rels");

    const { emb: embRelMap, media: mediaRelMap } = buildRelMaps(relBuf.toString("utf8"));

    // 1) Convert embeddings (OLE) -> MathML -> LaTeX (parallel with limit)
    const latexByRid = {};
    const mathmlByRid = {};
    const embEntries = Object.entries(embRelMap);
    await Promise.all(embEntries.map(async ([rid, embPath]) => {
      const f = (zip.files || []).find(x => x.path === embPath);
      if (!f) return;
      try {
        const buf = await f.buffer();
        const mml = await rubyConvertWithCache(buf, embPath);
        if (!mml) return;
        mathmlByRid[rid] = mml;
        const latex = mathmlToLatexSafe(mml);
        if (latex) latexByRid[rid] = latex;
      } catch (e) {
        // ignore per-item failures
      }
    }));

    // 2) Convert media -> dataUris (handle EMF/WMF)
    const imageByRid = {};
    const mediaEntries = Object.entries(mediaRelMap);
    await Promise.all(mediaEntries.map(async ([rid, mediaPath]) => {
      const f = (zip.files || []).find(x => x.path === mediaPath);
      if (!f) return;
      try {
        const buf = await f.buffer();
        const ext = getExtFromPath(mediaPath);
        if (ext === "emf" || ext === "wmf") {
          const png = convertEmfWmfToPng(buf, ext);
          if (png) {
            imageByRid[rid] = `data:image/png;base64,${png.toString("base64")}`;
            return;
          }
        }
        imageByRid[rid] = `data:${mimeFromExt(mediaPath)};base64,${buf.toString("base64")}`;
      } catch (e) {}
    }));

    // 3) Build inlineHtml using preserve-order renderer
    const debug = {
      embeddings: Object.keys(embRelMap).length,
      latexCount: Object.keys(latexByRid).length,
      imagesCount: Object.keys(imageByRid).length,
      imagesInjected: 0,
      seenOleRuns: 0,
      seenOle: 0,
      oleInjected: 0,
      ignoredRids: 0,
      sampleRids: [],
      mathmlByRidCount: Object.keys(mathmlByRid).length
    };
    const ctx = { latexByRid, imageByRid, debug };

    let inlineHtml = buildInlineHtml(docBuf.toString("utf8"), ctx);
    inlineHtml = formatExamLayout(inlineHtml);
    inlineHtml = removeUnsupportedImages(inlineHtml);

    // 4) Parse exam from inlineHtml (preserve underline)
    const exam = parseExamFromInlineHtml(inlineHtml) || { version: 8, questions: [], sections: [] };

    // 5) Attach section order to questions and build blocks (section + question sequence)
    attachSectionOrderToQuestions(exam, exam.sections || []);
    const blocks = buildOrderedBlocksFromExam(exam);

    // 6) Return response (compatible with original server)
    res.json({
      ok: true,
      total: exam.questions.length,
      sections: exam.sections || [],
      blocks,
      exam,
      latex: latexByRid,
      images: imageByRid,
      rawInlineHtml: inlineHtml,
      debug,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

app.get("/ping", (_, res) => res.send("ok"));
app.get("/health", (_, res) => res.json({ ok: true, node: process.version }));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
