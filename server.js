// server.js
// ✅ FULL CODE (FIX: PHẦN đúng vị trí + FIX mất căn thức \sqrt + FIX MathML prefix m:)
// - Server trả `blocks` đã trộn section + question đúng thứ tự (frontend render chuẩn như Word)
// - MathType OLE -> MathML -> LaTeX
// - FIX: MathML namespaced tags <m:msqrt> / <m:math> => strip prefix để converter nhận đúng
// - FIX: menclose notation="radical" => msqrt
// - FIX: mo √ ... => msqrt
// - ✅ HARD FIX: Tokenize <msqrt>/<mroot> BEFORE convert, then restore => gần như không thể rơi \sqrt
// - ✅ Stability: giới hạn song song khi convert OLE để tránh quá tải (MATH_CONCURRENCY)
//
// Chạy: node server.js
// Yêu cầu: inkscape (convert emf/wmf), ruby + mt2mml.rb (fallback MathType)
// npm i express multer unzipper cors mathml-to-latex
//
// ENV:
// - PORT (default 3000)
// - MATH_CONCURRENCY (default 3)

import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import crypto from "crypto";
import { execFile, execFileSync } from "child_process";
import { MathMLToLaTeX } from "mathml-to-latex";

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

function sha1(buf) {
  return crypto.createHash("sha1").update(buf).digest("hex");
}

/* ================= Simple concurrency limiter ================= */

function createLimiter(max = 3) {
  let active = 0;
  const q = [];
  const next = () => {
    if (active >= max) return;
    const job = q.shift();
    if (!job) return;
    active++;
    Promise.resolve()
      .then(job.fn)
      .then(job.resolve, job.reject)
      .finally(() => {
        active--;
        next();
      });
  };
  return function limit(fn) {
    return new Promise((resolve, reject) => {
      q.push({ fn, resolve, reject });
      next();
    });
  };
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

/* ================= MathType OLE -> MathML -> LaTeX ================= */

function extractMathMLFromOleScan(buf) {
  // try utf8 <math>
  const utf8 = buf.toString("utf8");
  let i = utf8.indexOf("<math");
  if (i !== -1) {
    let j = utf8.indexOf("</math>", i);
    if (j !== -1) return utf8.slice(i, j + 7);
  }

  // try utf8 namespaced <m:math>
  i = utf8.indexOf("<m:math");
  if (i !== -1) {
    let j = utf8.indexOf("</m:math>", i);
    if (j !== -1) return utf8.slice(i, j + 9);
  }

  // try utf16le <math>
  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    let j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }

  // try utf16le <m:math>
  i = u16.indexOf("<m:math");
  if (i !== -1) {
    let j = u16.indexOf("</m:math>", i);
    if (j !== -1) return u16.slice(i, j + 9);
  }

  return null;
}

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

/* ================== MathML preprocess ================== */

const SQRT_MATHML_RE =
  /(msqrt|mroot|menclose|radical|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

function stripMathMLTagPrefixes(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);
  // Strip common namespace prefixes in tags: m:, mml:, math:
  s = s.replace(/<\s*\/\s*(m|mml|math)\s*:/gi, "</");
  s = s.replace(/<\s*(m|mml|math)\s*:/gi, "<");
  return s;
}

function ensureMathMLNamespace(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml).replace(/<\?xml[^>]*\?>/gi, "").trim();
  // only add xmlns for <math ...> if missing
  s = s.replace(
    /<math(?![^>]*\bxmlns=)/i,
    '<math xmlns="http://www.w3.org/1998/Math/MathML"'
  );
  return s;
}

function normalizeMtable(mathml) {
  if (!mathml) return mathml;
  return String(mathml).replace(/<mtable\b[^>]*>/gi, "<mtable>");
}

// menclose radical -> msqrt + mo √ -> msqrt
function preprocessMathMLForSqrt(mathml) {
  if (!mathml) return mathml;
  let s = stripMathMLTagPrefixes(String(mathml));

  // 1) menclose radical -> msqrt
  s = s.replace(
    /<menclose\b([^>]*)\bnotation\s*=\s*["']radical["']([^>]*)>([\s\S]*?)<\/menclose>/gi,
    "<msqrt>$3</msqrt>"
  );

  // 2) <mo>√</mo> + next node -> msqrt
  const moSqrt =
    String.raw`<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>`;

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
    new RegExp(
      moSqrt + String.raw`\s*<mfenced([^>]*)>([\s\S]*?)<\/mfenced>`,
      "gi"
    ),
    "<msqrt><mfenced$1>$2</mfenced></msqrt>"
  );

  return s;
}

/* ================== LATEX POSTPROCESS + SQRT FIX ================== */

function sanitizeLatexStrict(latex) {
  if (!latex) return latex;
  latex = String(latex).replace(/\s+/g, " ").trim();

  latex = latex
    .replace(
      /\\left(?!\s*(\(|\[|\\\{|\\langle|\\vert|\\\||\||\.))/g,
      ""
    )
    .replace(
      /\\right(?!\s*(\)|\]|\\\}|\\rangle|\\vert|\\\||\||\.))/g,
      ""
    );

  const tokens = latex.match(/\\left\b|\\right\b/g) || [];
  let bal = 0;
  let broken = false;
  for (const t of tokens) {
    if (t === "\\left") bal++;
    else {
      if (bal === 0) {
        broken = true;
        break;
      }
      bal--;
    }
  }
  if (bal !== 0) broken = true;

  if (broken) latex = latex.replace(/\\left\s*/g, "").replace(/\\right\s*/g, "");
  return latex;
}

function fixSetBracesHard(latex) {
  let s = String(latex || "");

  s = s.replace(
    /\\underset\s*\{([^}]*)\}\s*\{\s*l\s*i\s*m\s*\}/gi,
    "\\underset{$1}{\\lim}"
  );
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

  if (
    (s.includes("\\backslash \\{") || s.includes("\\setminus \\{")) &&
    !s.includes("\\}")
  ) {
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

function fixPiecewiseFunction(latex) {
  let s = String(latex || "");

  s = s.replace(/\(\.\s+/g, "(");
  s = s.replace(/\s+\.\)/g, ")");
  s = s.replace(/\[\.\s+/g, "[");
  s = s.replace(/\s+\.\]/g, "]");

  const piecewiseMatch = s.match(/(?<!\\)\{\.\s+/);
  if (piecewiseMatch) {
    const startIdx = piecewiseMatch.index;
    const contentStart = startIdx + piecewiseMatch[0].length;

    let braceCount = 1;
    let endIdx = contentStart;
    let foundEnd = false;

    for (let i = contentStart; i < s.length; i++) {
      const ch = s[i];
      const prevCh = i > 0 ? s[i - 1] : "";
      if (prevCh === "\\") continue;

      if (ch === "{") braceCount++;
      else if (ch === "}") {
        braceCount--;
        if (braceCount === 0) {
          endIdx = i;
          foundEnd = true;
          break;
        }
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

function postprocessLatexSqrt(latex) {
  if (!latex) return latex;
  let s = String(latex);

  // Some converters output \surd instead of \sqrt
  s = s.replace(/\\surd\b/g, "\\sqrt{}");

  // literal sqrt remained
  s = s.replace(/√\s*\{([^}]+)\}/g, "\\sqrt{$1}");
  s = s.replace(/√\s*\(([^)]+)\)/g, "\\sqrt{$1}");
  s = s.replace(/√\s*(\d+)/g, "\\sqrt{$1}");
  s = s.replace(/√\s*([a-zA-Z])/g, "\\sqrt{$1}");

  // \sqrt without braces
  s = s.replace(/\\sqrt\s+(\d+)(?![}\d])/g, "\\sqrt{$1}");
  s = s.replace(/\\sqrt\s+([a-zA-Z])(?![}\w])/g, "\\sqrt{$1}");

  // empty sqrt
  s = s.replace(/\\sqrt\s*\{\s*\}/g, "\\sqrt{\\phantom{x}}");

  // malformed spaces
  s = s.replace(/\\sqrt\s+\{/g, "\\sqrt{");

  // nth root \root{n}\of{x} -> \sqrt[n]{x}
  s = s.replace(
    /\\root\s*\{([^}]+)\}\s*\\of\s*\{([^}]+)\}/g,
    "\\sqrt[$1]{$2}"
  );
  s = s.replace(/\\sqrt\s*\[\s*(\d+)\s*\]\s*\{/g, "\\sqrt[$1]{");

  return s;
}

/**
 * ✅ HARD WRAP:
 * If MathML had sqrt/root but LaTeX output has no \sqrt => wrap result in \sqrt{...}
 */
function fixSqrtLatex(latex, mathmlMaybe = "") {
  let s = String(latex || "").trim();
  const mml = stripMathMLTagPrefixes(String(mathmlMaybe || ""));

  const mmlHasSqrt =
    SQRT_MATHML_RE.test(mml) ||
    /<msqrt\b/i.test(mml) ||
    /<mroot\b/i.test(mml) ||
    /<menclose\b[^>]*notation\s*=\s*["']radical["']/i.test(mml);

  const latexHasSqrt = /\\sqrt\b|\\root\b/.test(s);

  if (mmlHasSqrt && !latexHasSqrt) {
    if (!s) return "\\sqrt{}";
    return `\\sqrt{${s}}`;
  }

  return s;
}

function postProcessLatex(latex, mathmlMaybe = "") {
  let s = latex || "";
  s = sanitizeLatexStrict(s);
  s = fixSetBracesHard(s);
  s = restoreArrowAndCoreCommands(s);
  s = fixPiecewiseFunction(s);

  // soft sqrt fixes
  s = postprocessLatexSqrt(s);

  // hard wrap last
  s = fixSqrtLatex(s, mathmlMaybe);

  return String(s || "").replace(/\s+/g, " ").trim();
}

/* ================= ✅ SQRT TOKEN CONVERTER (KHÓ RƠI CĂN NHẤT) ================= */

/** Find balanced tag blocks <tag ...> ... </tag> */
function extractBalancedTagBlocks(xml, tagName) {
  const blocks = [];
  const openRe = new RegExp(`<${tagName}\\b[^>]*>`, "ig");
  const closeRe = new RegExp(`</${tagName}>`, "ig");

  let m;
  while ((m = openRe.exec(xml)) !== null) {
    const start = m.index;
    const openLen = m[0].length;

    let depth = 1;
    let i = start + openLen;

    while (depth > 0) {
      const nextOpen = openRe.exec(xml);
      const nextClose = closeRe.exec(xml);

      const o = nextOpen ? nextOpen.index : Infinity;
      const c = nextClose ? nextClose.index : Infinity;

      if (c === Infinity) break; // malformed

      if (o < c) {
        depth++;
        i = o + (nextOpen[0] ? nextOpen[0].length : 0);
      } else {
        depth--;
        i = c + (nextClose[0] ? nextClose[0].length : 0);
      }
    }

    const end = i;
    if (end > start) {
      blocks.push({ start, end, xml: xml.slice(start, end) });
      // reset regex lastIndex to continue AFTER this block
      openRe.lastIndex = start + 1;
      closeRe.lastIndex = start + 1;
    }
  }

  blocks.sort((a, b) => a.start - b.start);
  return blocks;
}

/** Split direct children inside <mroot> ... </mroot> into [base, index] (heuristic good for mt2mml output) */
function splitMrootChildren(innerXml) {
  const s = innerXml.trim();
  let depth = 0;
  let cut = -1;

  for (let i = 0; i < s.length; i++) {
    if (s[i] === "<") {
      if (s.slice(i, i + 2) === "</") depth--;
      else if (s.slice(i, i + 5).toLowerCase() === "<mrow") depth++;
      else if (s.slice(i, i + 3).toLowerCase() === "<mi") depth++;
      else if (s.slice(i, i + 3).toLowerCase() === "<mn") depth++;
      else if (s.slice(i, i + 3).toLowerCase() === "<mo") depth++;
      else if (s.slice(i, i + 6).toLowerCase() === "<msqrt") depth++;
      else if (s.slice(i, i + 6).toLowerCase() === "<mroot") depth++;
      else if (s.slice(i, i + 9).toLowerCase() === "<menclose") depth++;
    }

    // boundary when first element closed at top level
    if (depth === 0 && i > 0) {
      const maybe = s.slice(0, i + 1);
      if (/<\/\w+>\s*$/.test(maybe)) {
        cut = i + 1;
        break;
      }
    }
  }

  if (cut < 0) return [s, ""];
  const a = s.slice(0, cut).trim();
  const b = s.slice(cut).trim();
  return [a, b];
}

/**
 * Replace all <msqrt>/<mroot> blocks by tokens BEFORE MathMLToLaTeX.convert,
 * then restore tokens into \sqrt{...} / \sqrt[n]{...}
 */
function convertMathMLWithSqrtTokens(mathml) {
  let mm = stripMathMLTagPrefixes(String(mathml || "")).trim();

  // wrap if body-only
  const hasMathTag = /<\s*math\b/i.test(mm);
  const looksLikeBody =
    /<(mrow|mi|mn|mo|msqrt|mroot|mfrac|msup|msub|msubsup|menclose)\b/i.test(mm);
  if (!hasMathTag && looksLikeBody) {
    mm = `<math xmlns="http://www.w3.org/1998/Math/MathML">${mm}</math>`;
  }
  if (!/<\s*math\b/i.test(mm)) return "";

  mm = ensureMathMLNamespace(mm);
  mm = normalizeMtable(mm);
  mm = preprocessMathMLForSqrt(mm);

  // extract blocks
  const sqrtBlocks = extractBalancedTagBlocks(mm, "msqrt");
  const rootBlocks = extractBalancedTagBlocks(mm, "mroot");

  const tokenMap = []; // { token, type, xml }
  let replaced = mm;

  // replace from back to front to keep indices valid
  const all = [
    ...sqrtBlocks.map((b) => ({ ...b, type: "msqrt" })),
    ...rootBlocks.map((b) => ({ ...b, type: "mroot" })),
  ].sort((a, b) => b.start - a.start);

  let k = 0;
  for (const b of all) {
    const token = `__SQRT_TOKEN_${++k}__`;
    tokenMap.push({ token, type: b.type, xml: b.xml });

    // replace whole block with a harmless identifier node
    replaced =
      replaced.slice(0, b.start) + `<mi>${token}</mi>` + replaced.slice(b.end);
  }

  // convert the big expression (sqrt removed)
  let latexMain = "";
  try {
    latexMain = (MathMLToLaTeX.convert(replaced) || "").trim();
  } catch {
    latexMain = "";
  }

  // restore each token with manual sqrt/root conversion
  for (const item of tokenMap) {
    let latexRep = "";

    if (item.type === "msqrt") {
      const inner = item.xml
        .replace(/^<msqrt\b[^>]*>/i, "")
        .replace(/<\/msqrt>\s*$/i, "");
      const innerLatex = convertMathMLWithSqrtTokens(
        `<math xmlns="http://www.w3.org/1998/Math/MathML">${inner}</math>`
      );
      latexRep = `\\sqrt{${innerLatex || ""}}`;
    } else {
      const inner = item.xml
        .replace(/^<mroot\b[^>]*>/i, "")
        .replace(/<\/mroot>\s*$/i, "");
      const [baseXml, indexXml] = splitMrootChildren(inner);

      const baseLatex = convertMathMLWithSqrtTokens(
        `<math xmlns="http://www.w3.org/1998/Math/MathML">${baseXml}</math>`
      );
      const idxLatex = convertMathMLWithSqrtTokens(
        `<math xmlns="http://www.w3.org/1998/Math/MathML">${indexXml}</math>`
      );
      latexRep = `\\sqrt[${idxLatex || ""}]{${baseLatex || ""}}`;
    }

    // replace token in latex output (token was in <mi>)
    const re = new RegExp(
      item.token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      "g"
    );
    latexMain = latexMain.replace(re, latexRep);
  }

  return latexMain.trim();
}

/* ================= ✅ Safe MathML -> LaTeX (uses token converter) ================= */

/**
 * Accept MathML:
 * - <math ...>...</math>
 * - <m:math ...>...</m:math>   (we strip prefix)
 * - or only body (<mrow>...</mrow>, <msqrt>...) => auto-wrap into <math>
 */
function mathmlToLatexSafe(mml) {
  try {
    if (!mml) return "";
    let mm = String(mml).trim();

    mm = stripMathMLTagPrefixes(mm);

    // wrap body-only
    const hasMathTag = /<\s*math\b/i.test(mm);
    const looksLikeMathMLBody =
      /<(mrow|mi|mn|mo|msqrt|mroot|mfrac|msup|msub|msubsup|menclose)\b/i.test(mm);

    if (!hasMathTag && looksLikeMathMLBody) {
      mm = `<math xmlns="http://www.w3.org/1998/Math/MathML">${mm}</math>`;
    }

    if (!/<\s*math\b/i.test(mm)) return "";

    mm = ensureMathMLNamespace(mm);
    mm = normalizeMtable(mm);
    mm = preprocessMathMLForSqrt(mm);

    // ✅ Token-based convert (avoid losing sqrt)
    const latex0 = convertMathMLWithSqrtTokens(mm);

    // Postprocess + final hard wrap
    return postProcessLatex(latex0 || "", mm);
  } catch {
    return "";
  }
}

/* ================= MathType FIRST ================= */

async function tokenizeMathTypeOleFirst(docXml, rels, zipFiles, images) {
  let idx = 0;
  const found = {}; // key -> { oleTarget, previewRid, oleHash }
  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;

  docXml = docXml.replace(OBJECT_RE, (block) => {
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid);
    if (!oleTarget) return block;

    const vmlRid = block.match(
      /<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/
    );
    const blipRid = block.match(
      /<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/
    );
    const previewRid = vmlRid?.[1] || blipRid?.[1] || null;

    const key = `mathtype_${++idx}`;
    found[key] = { oleTarget, previewRid, oleHash: null };
    return `[!m:$${key}$]`;
  });

  const latexMap = {};
  const oleCache = new Map(); // oleHash -> latex
  const mmlCache = new Map(); // oleHash -> mml
  const limit = createLimiter(
    Math.max(1, Number(process.env.MATH_CONCURRENCY || 3))
  );

  await Promise.all(
    Object.entries(found).map(([key, info]) =>
      limit(async () => {
        const oleFull = normalizeTargetToWordPath(info.oleTarget);
        const oleBuf = await getZipEntryBuffer(zipFiles, oleFull);

        // No OLE => fallback preview image
        if (!oleBuf) {
          latexMap[key] = "";
          return;
        }

        const h = sha1(oleBuf);
        info.oleHash = h;

        // Cache latex by hash
        if (oleCache.has(h)) {
          latexMap[key] = oleCache.get(h) || "";
          return;
        }

        let mml = mmlCache.get(h) || "";
        if (!mml) mml = extractMathMLFromOleScan(oleBuf) || "";

        if (!mml) {
          try {
            mml = await rubyOleToMathML(oleBuf);
          } catch {
            mml = "";
          }
        }
        if (mml) mmlCache.set(h, mml);

        const latex = mml ? mathmlToLatexSafe(mml) : "";
        if (latex) {
          latexMap[key] = latex;
          oleCache.set(h, latex);
          return;
        }

        // fallback preview image if conversion failed
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
                    images[`fallback_${key}`] = `data:image/png;base64,${pngBuf.toString(
                      "base64"
                    )}`;
                    latexMap[key] = "";
                    oleCache.set(h, "");
                    return;
                  }
                } catch {}
              }
              images[`fallback_${key}`] = `data:${mime};base64,${imgBuf.toString(
                "base64"
              )}`;
            }
          }
        }

        latexMap[key] = "";
        oleCache.set(h, "");
      })
    )
  );

  return { outXml: docXml, latexMap };
}

/* ================= Images AFTER MathType ================= */

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

/* ================= Text (GIỮ token + underline) ================= */

function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // Protect tokens BEFORE stripping tags
  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

  // Convert each run while preserving underline
  x = x.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
    const hasU =
      /<w:u\b[^>]*\/>/.test(run) &&
      !/<w:u\b[^>]*w:val="none"[^>]*\/>/.test(run);

    let inner = run.replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g, "");
    inner = inner.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (_, t) => t ?? "");
    inner = inner.replace(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g,
      (_, t) => t ?? ""
    );

    inner = inner.replace(/<[^>]+>/g, "");
    if (!inner) return "";
    return hasU ? `<u>${inner}</u>` : inner;
  });

  // Remove remaining tags outside runs, but keep <u>
  x = x.replace(/<(?!\/?u\b)[^>]+>/g, "");

  // Restore tokens stable form
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

/* ================= SECTION TITLES (PHẦN ...) ================= */

function extractSectionTitles(rawText) {
  const text = String(rawText || "").replace(/\r/g, "");

  // anchors "Câu X."
  const qRe = /(^|\n)\s*Câu\s+(\d+)\./gi;
  const qAnchors = [];
  let qm;
  while ((qm = qRe.exec(text)) !== null) {
    qAnchors.push({
      idx: qm.index + (qm[1] ? qm[1].length : 0),
      no: Number(qm[2]),
    });
  }

  // anchors "PHẦN ..."
  const sRe =
    /(^|\n)\s*(?:[-•–]\s*)?PHẦN\s+([0-9]+|[IVXLCDM]+)\s*[\.\:\-]?\s*([^\n]*)/gi;

  const sections = [];
  let sm;
  while ((sm = sRe.exec(text)) !== null) {
    const startChar = sm.index + (sm[1] ? sm[1].length : 0);

    let titleLine = `PHẦN ${sm[2]}. ${sm[3] || ""}`.trim();
    const cutIdx = titleLine.search(/(?=\bCâu\s+\d+\.)/i);
    if (cutIdx > 0) titleLine = titleLine.slice(0, cutIdx).trim();

    sections.push({
      title: titleLine,
      order: sections.length + 1,
      startChar,
      endChar: null,
      firstQuestionNo: null,
      questionCount: 0,
      questionIndexStart: null,
      questionIndexEnd: null,
    });
  }

  for (let i = 0; i < sections.length; i++) {
    sections[i].endChar =
      i + 1 < sections.length ? sections[i + 1].startChar : text.length;
  }

  for (const sec of sections) {
    const startIdx = qAnchors.findIndex(
      (q) => q.idx >= sec.startChar && q.idx < sec.endChar
    );
    if (startIdx === -1) continue;

    let endIdx = qAnchors.length;
    for (let k = startIdx; k < qAnchors.length; k++) {
      if (qAnchors[k].idx >= sec.endChar) {
        endIdx = k;
        break;
      }
    }

    sec.questionIndexStart = startIdx;
    sec.questionIndexEnd = endIdx;
    sec.questionCount = endIdx - startIdx;
    sec.firstQuestionNo = qAnchors[startIdx]?.no ?? null;
  }

  return sections;
}

/* ================== EXAM PARSER (GIỮ NGUYÊN) ================== */

function stripTagsToPlain(s) {
  return String(s || "")
    .replace(/<u[^>]*>/gi, "")
    .replace(/<\/u>/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function detectHasMCQ(plain) {
  const marks = plain.match(/\b[ABCD]\./g) || [];
  return new Set(marks).size >= 2;
}

function detectHasTF4(plain) {
  const marks = plain.match(/\b[a-d]\)/gi) || [];
  return new Set(marks.map((x) => x.toLowerCase())).size >= 2;
}

function extractUnderlinedKeys(blockText) {
  const keys = { mcq: null, tf: [] };
  const s = String(blockText || "");

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

function normalizeUnderlinedMarkersForSplit(s) {
  let x = String(s || "");
  x = x.replace(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./gi, "$1.");
  x = x.replace(/<u[^>]*>\s*([A-D])\.\s*<\/u>/gi, "$1.");
  x = x.replace(/<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi, "$1)");
  x = x.replace(/<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi, "$1)");
  return x;
}

function findSolutionMarkerIndex(text, fromIndex = 0) {
  const s = String(text || "");
  const re = /(Lời\s*giải|Giải\s*chi\s*tiết|Hướng\s*dẫn\s*giải)/i;
  const sub = s.slice(fromIndex);
  const m = re.exec(sub);
  if (!m) return -1;
  return fromIndex + m.index;
}

function splitSolutionSections(tailText) {
  let s = String(tailText || "").trim();
  if (!s) return { solution: "", detail: "" };

  const reCT = /(Giải\s*chi\s*tiết)/i;
  const matchCT = reCT.exec(s);
  if (matchCT) {
    const idxCT = matchCT.index;
    return {
      solution: s.slice(0, idxCT).trim(),
      detail: s.slice(idxCT).trim(),
    };
  }
  return { solution: s, detail: "" };
}

function cleanStemFromQuestionNo(s) {
  return String(s || "").replace(/^Câu\s+\d+\.?\s*/i, "").trim();
}

function splitChoicesTextABCD(blockText) {
  let s = normalizeUnderlinedMarkersForSplit(blockText);
  s = s.replace(/\r/g, "");

  const solIdx = findSolutionMarkerIndex(s, 0);
  const main = solIdx >= 0 ? s.slice(0, solIdx) : s;
  const tail = solIdx >= 0 ? s.slice(solIdx) : "";

  const re = /(^|\n)\s*(\*?)([A-D])\.\s*/g;

  const hits = [];
  let m;
  while ((m = re.exec(main)) !== null) {
    hits.push({ idx: m.index + m[1].length, star: m[2] === "*", key: m[3] });
  }
  if (hits.length < 2) return null;

  const out = {
    stem: main.slice(0, hits[0].idx).trim(),
    choices: { A: "", B: "", C: "", D: "" },
    starredCorrect: null,
    tail,
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : main.length;
    let seg = main.slice(start, end).trim();
    seg = seg.replace(/^(\*?)([A-D])\.\s*/i, "");
    out.choices[key] = seg.trim();
    if (hits[i].star) out.starredCorrect = key;
  }
  return out;
}

function splitStatementsTextabcd(blockText) {
  let s = normalizeUnderlinedMarkersForSplit(blockText);
  s = s.replace(/\r/g, "");

  const solIdx = findSolutionMarkerIndex(s, 0);
  const main = solIdx >= 0 ? s.slice(0, solIdx) : s;
  const tail = solIdx >= 0 ? s.slice(solIdx) : "";

  const re = /(^|\n)\s*([a-d])\)\s*/gi;
  const hits = [];
  let m;
  while ((m = re.exec(main)) !== null) {
    hits.push({ idx: m.index + m[1].length, key: m[2].toLowerCase() });
  }
  if (hits.length < 2) return null;

  const out = {
    stem: main.slice(0, hits[0].idx).trim(),
    statements: { a: "", b: "", c: "", d: "" },
    tail,
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : main.length;
    let seg = main.slice(start, end).trim();
    seg = seg.replace(/^([a-d])\)\s*/i, "");
    out.statements[key] = seg.trim();
  }
  return out;
}

function parseExamFromText(text) {
  const blocks = String(text || "").split(/(?=Câu\s+\d+\.)/);
  const exam = { version: 10, questions: [] };

  for (const block of blocks) {
    if (!/^Câu\s+\d+\./i.test(block)) continue;

    const qnoMatch = block.match(/^Câu\s+(\d+)\./i);
    const no = qnoMatch ? Number(qnoMatch[1]) : null;

    const under = extractUnderlinedKeys(block);
    const plain = stripTagsToPlain(block);

    const isMCQ = detectHasMCQ(plain);
    const isTF4 = !isMCQ && detectHasTF4(plain);

    if (isMCQ) {
      const parts = splitChoicesTextABCD(block);
      const tail = parts?.tail || "";
      const solParts = splitSolutionSections(tail);

      const answer = parts?.starredCorrect || under.mcq || null;

      exam.questions.push({
        no,
        type: "mcq",
        stem: cleanStemFromQuestionNo(parts?.stem || block),
        choices: {
          A: parts?.choices?.A || "",
          B: parts?.choices?.B || "",
          C: parts?.choices?.C || "",
          D: parts?.choices?.D || "",
        },
        answer,
        solution: solParts.solution || "",
        detail: solParts.detail || "",
        _plain: plain,
      });
      continue;
    }

    if (isTF4) {
      const parts = splitStatementsTextabcd(block);
      const tail = parts?.tail || "";
      const solParts = splitSolutionSections(tail);

      const ans = { a: null, b: null, c: null, d: null };
      for (const k of ["a", "b", "c", "d"]) {
        if (under.tf.includes(k)) ans[k] = true;
      }

      exam.questions.push({
        no,
        type: "tf4",
        stem: cleanStemFromQuestionNo(parts?.stem || block),
        statements: {
          a: parts?.statements?.a || "",
          b: parts?.statements?.b || "",
          c: parts?.statements?.c || "",
          d: parts?.statements?.d || "",
        },
        answer: ans,
        solution: solParts.solution || "",
        detail: solParts.detail || "",
        _plain: plain,
      });
      continue;
    }

    const solIdx = findSolutionMarkerIndex(block, 0);
    const stemPart = solIdx >= 0 ? block.slice(0, solIdx).trim() : block.trim();
    const tailPart = solIdx >= 0 ? block.slice(solIdx).trim() : "";

    const solParts = splitSolutionSections(tailPart);

    exam.questions.push({
      no,
      type: "short",
      stem: cleanStemFromQuestionNo(stemPart),
      boxes: 4,
      solution: solParts.solution || tailPart || "",
      detail: solParts.detail || "",
      _plain: plain,
    });
  }

  return exam;
}

function legacyQuestionsFromExam(exam) {
  const out = [];
  for (const q of exam.questions) {
    if (q.type !== "mcq") continue;
    out.push({
      type: "multiple_choice",
      content: q.stem,
      choices: [
        { label: "A", text: q.choices.A },
        { label: "B", text: q.choices.B },
        { label: "C", text: q.choices.C },
        { label: "D", text: q.choices.D },
      ],
      correct: q.answer,
      solution: [q.solution, q.detail].filter(Boolean).join("\n").trim(),
    });
  }
  return out;
}

/* ================= helper: gán sectionOrder cho từng question ================= */

function attachSectionOrderToQuestions(exam, sections) {
  if (!exam?.questions?.length || !Array.isArray(sections)) return;

  for (const q of exam.questions) {
    q.sectionOrder = null;
    q.sectionTitle = null;
  }

  for (const sec of sections) {
    if (
      typeof sec.questionIndexStart !== "number" ||
      typeof sec.questionIndexEnd !== "number"
    ) {
      continue;
    }
    const a = Math.max(0, sec.questionIndexStart);
    const b = Math.min(exam.questions.length, sec.questionIndexEnd);
    for (let i = a; i < b; i++) {
      exam.questions[i].sectionOrder = sec.order;
      exam.questions[i].sectionTitle = sec.title;
    }
  }
}

/* ================= ✅ blocks trộn section + question đúng thứ tự ================= */

function buildOrderedBlocks(exam) {
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

    // 1) MathType -> LaTeX (and fallback images)
    const images = {};
    const mt = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mt.outXml;
    const latexMap = mt.latexMap;

    // 2) normal images
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // 3) text (giữ token + underline)
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) parse exam output
    const exam = parseExamFromText(text);

    // sections theo vị trí + index câu toàn cục
    const sections = extractSectionTitles(text);
    exam.sections = sections;

    // gán sectionOrder/sectionTitle cho từng question
    attachSectionOrderToQuestions(exam, sections);

    // ✅ blocks đã trộn đúng thứ tự để frontend render chuẩn như Word
    const blocks = buildOrderedBlocks(exam);

    // legacy (để tương thích)
    const questions = legacyQuestionsFromExam(exam);

    res.json({
      ok: true,
      total: exam.questions.length,
      sections,
      blocks,
      exam,
      questions,
      latex: latexMap,
      images,
      rawText: text,
      debug: {
        latexCount: Object.keys(latexMap).length,
        imagesCount: Object.keys(images).length,
        mathConcurrency: Math.max(1, Number(process.env.MATH_CONCURRENCY || 3)),
        exam: {
          questions: exam.questions.length,
          mcq: exam.questions.filter((x) => x.type === "mcq").length,
          tf4: exam.questions.filter((x) => x.type === "tf4").length,
          short: exam.questions.filter((x) => x.type === "short").length,
        },
      },
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
