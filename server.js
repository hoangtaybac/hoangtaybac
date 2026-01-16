import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
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

/* ================== LATEX POSTPROCESS (GHÉP XỬ LÝ CĂN + CASES) ================== */

const SQRT_MATHML_RE = /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

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

/**
 * Fix piecewise / cases (hệ phương trình, hàm phân đoạn)
 */
function fixPiecewiseFunction(latex) {
  let s = String(latex || "");

  // Fix broken parentheses/brackets like "(. " "[. "
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
    // normalize new rows
    content = content.replace(/\s+\\\s+(?=\d)/g, " \\\\ ");

    const before = s.slice(0, startIdx);
    const after = foundEnd ? s.slice(endIdx + 1) : "";
    s = before + `\\begin{cases} ${content} \\end{cases}` + after;
  }

  return s;
}

/**
 * ✅ XỬ LÝ CĂN (sqrt) “cứng”
 */
function fixSqrtLatex(latex, mathmlMaybe = "") {
  let s = String(latex || "");

  // √(x+1) => \sqrt{x+1}
  s = s.replace(/√\s*\(\s*([\s\S]*?)\s*\)/g, "\\sqrt{$1}");
  // √x => \sqrt{x}
  s = s.replace(/√\s*([A-Za-z0-9]+)\b/g, "\\sqrt{$1}");

  if (SQRT_MATHML_RE.test(String(mathmlMaybe || ""))) {
    const hasSqrt = /\\sqrt\b|\\root\b/.test(s);
    if (!hasSqrt && s) {
      s = s.replace(/\bradic\b/gi, "\\sqrt{}");
    }
  }

  return s;
}

function postProcessLatex(latex, mathmlMaybe = "") {
  let s = latex || "";
  s = sanitizeLatexStrict(s);
  s = fixSetBracesHard(s);
  s = restoreArrowAndCoreCommands(s);
  s = fixPiecewiseFunction(s);
  s = fixSqrtLatex(s, mathmlMaybe);
  return String(s || "").replace(/\s+/g, " ").trim();
}

function mathmlToLatexSafe(mml) {
  try {
    if (!mml || !mml.includes("<math")) return "";
    const latex0 = (MathMLToLaTeX.convert(mml) || "").trim();
    return postProcessLatex(latex0, mml);
  } catch {
    return "";
  }
}

/**
 * MathType FIRST
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

    const vmlRid = block.match(
      /<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/
    );
    const blipRid = block.match(
      /<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/
    );
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

      // 1) scan MathML inside OLE
      let mml = "";
      if (oleBuf) mml = extractMathMLFromOleScan(oleBuf) || "";

      // 2) fallback ruby
      if (!mml && oleBuf) {
        try {
          mml = await rubyOleToMathML(oleBuf);
        } catch {
          mml = "";
        }
      }

      // 3) MathML -> LaTeX
      const latex = mml ? mathmlToLatexSafe(mml) : "";
      if (latex) {
        latexMap[key] = latex;
        return;
      }

      // 4) fallback preview image
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
    })
  );

  return { outXml: docXml, latexMap };
}

/**
 * Tokenize normal images AFTER MathType
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

/* ================= Text (GIỮ token + underline) ================= */

function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // 1) Protect tokens BEFORE stripping tags
  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

  // 2) Convert each run <w:r> while preserving underline + token text
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

/* ================= SECTION TITLES (PHẦN ...) ✅ BỔ SUNG ================= */

/**
 * Lấy các tiêu đề kiểu:
 * - "PHẦN 1. ..."
 * - "PHẦN 2: ..."
 * - "PHẦN I. ..." (roman)
 * Trả về theo thứ tự xuất hiện trong rawText.
 * ✅ Không thay đổi thuật toán parseExamFromText của bạn.
 */
function extractSectionTitles(rawText) {
  const s = String(rawText || "").replace(/\r/g, "");
  const lines = s.split("\n");

  const sections = [];
  const seen = new Set();

  const re = /^\s*PHẦN\s+([0-9]+|[IVXLCDM]+)\s*[\.\:\-]?\s*(.*)\s*$/i;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const m = re.exec(line);
    if (!m) continue;

    const title = line; // giữ nguyên đúng như rawText
    const key = title.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);

    // tìm câu đầu tiên phía sau tiêu đề (nếu có)
    let firstQuestionNo = null;
    for (let j = i + 1; j < Math.min(i + 80, lines.length); j++) {
      const q = lines[j].match(/^\s*Câu\s+(\d+)\./i);
      if (q) {
        firstQuestionNo = Number(q[1]);
        break;
      }
    }

    sections.push({
      title,
      order: sections.length + 1,
      firstQuestionNo,
      lineIndex: i,
    });
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
  const exam = { version: 9, questions: [] };

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

    // ✅ BỔ SUNG: lấy tiêu đề PHẦN từ rawText (KHÔNG thay đổi thuật toán parse câu)
    const sections = extractSectionTitles(text);

    // 4) parse exam output (mcq/tf4/short) — GIỮ NGUYÊN
    const exam = parseExamFromText(text);

    // ✅ BỔ SUNG: nhét sections vào exam (chỉ thêm field mới)
    exam.sections = sections;

    // 5) legacy questions output (mcq only)
    const questions = legacyQuestionsFromExam(exam);

    res.json({
      ok: true,
      total: exam.questions.length,

      // ✅ BỔ SUNG: sections ở top-level (dễ dùng cho front)
      sections,

      exam,
      questions,
      latex: latexMap,
      images,
      rawText: text,
      debug: {
        latexCount: Object.keys(latexMap).length,
        imagesCount: Object.keys(images).length,
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
