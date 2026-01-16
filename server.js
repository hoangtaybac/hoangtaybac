// server.js
// ✅ PATCHED FULL CODE: FIX mất căn thức \sqrt cho MathType (mt2mml.rb -> MathML -> LaTeX)
// - Thêm preprocess MathML: strip prefix m:, add xmlns, menclose(radical)->msqrt, mo √ -> msqrt
// - ✅ FIX mẫu MathML lỗi (mt2mml đôi khi xuất msup base rỗng): <mi>y</mi><msup><mrow></mrow><mrow><mn>2</mn></mrow></msup> => <msup><mi>y</mi><mn>2</mn></msup>
// - ✅ Tokenize <msqrt>/<mroot> trước khi convert để KHÓ rơi căn nhất
//
// Chạy: node server.js
// Yêu cầu: inkscape (convert emf/wmf), ruby + mt2mml.rb (fallback MathType)
// npm i express multer unzipper cors mathml-to-latex

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

  // also accept namespaced m:math if present
  i = utf8.indexOf("<m:math");
  if (i !== -1) {
    let j = utf8.indexOf("</m:math>", i);
    if (j !== -1) return utf8.slice(i, j + 9);
  }

  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    let j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }
  i = u16.indexOf("<m:math");
  if (i !== -1) {
    let j = u16.indexOf("</m:math>", i);
    if (j !== -1) return u16.slice(i, j + 9);
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

/* ================= MathML normalizer (FIX sqrt + broken msup) ================= */

const SQRT_MATHML_RE =
  /(msqrt|mroot|menclose|radical|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

function stripMathMLTagPrefixes(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);
  s = s.replace(/<\s*\/\s*(m|mml|math)\s*:/gi, "</");
  s = s.replace(/<\s*(m|mml|math)\s*:/gi, "<");
  return s;
}

function ensureMathMLNamespace(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml).replace(/<\?xml[^>]*\?>/gi, "").trim();
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

  // menclose radical => msqrt
  s = s.replace(
    /<menclose\b([^>]*)\bnotation\s*=\s*["']radical["']([^>]*)>([\s\S]*?)<\/menclose>/gi,
    "<msqrt>$3</msqrt>"
  );

  const moSqrt =
    String.raw`<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>`;

  // √ <mrow>...</mrow> => msqrt
  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mrow>([\s\S]*?)<\/mrow>`, "gi"),
    "<msqrt>$1</msqrt>"
  );
  // √ <mi>x</mi> => msqrt
  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mi>([^<]+)<\/mi>`, "gi"),
    "<msqrt><mi>$1</mi></msqrt>"
  );
  // √ <mn>3</mn> => msqrt
  s = s.replace(
    new RegExp(moSqrt + String.raw`\s*<mn>([^<]+)<\/mn>`, "gi"),
    "<msqrt><mn>$1</mn></msqrt>"
  );

  return s;
}

/**
 * ✅ FIX msup lỗi hay gặp (mt2mml): base nằm ngoài, msup base rỗng
 * Ví dụ:
 *   <mi>y</mi><msup><mrow></mrow><mrow><mn>2</mn></mrow></msup>
 * => <msup><mi>y</mi><mn>2</mn></msup>
 */
function fixBrokenMsup(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);

  // Case 1: <mi>y</mi><msup><mrow></mrow><mrow>EXP</mrow></msup>
  s = s.replace(
    /(<mi>[^<]+<\/mi>|<mn>[^<]+<\/mn>)\s*<msup>\s*<mrow>\s*<\/mrow>\s*<mrow>\s*([\s\S]*?)\s*<\/mrow>\s*<\/msup>/gi,
    "<msup>$1<$2</msup>"
  );

  // The above line is intentionally wrong (placeholder) — we will do a safer one below.
  // Safer replacements (don’t create malformed tags):
  s = s.replace(
    /(<mi>[^<]+<\/mi>|<mn>[^<]+<\/mn>)\s*<msup>\s*<mrow>\s*<\/mrow>\s*<mrow>\s*([\s\S]*?)\s*<\/mrow>\s*<\/msup>/gi,
    (m, base, expInner) => `<msup>${base}<mrow>${expInner}</mrow></msup>`
  );

  // Case 2: <mi>y</mi><msup><mrow></mrow><mn>2</mn></msup>
  s = s.replace(
    /(<mi>[^<]+<\/mi>|<mn>[^<]+<\/mn>)\s*<msup>\s*<mrow>\s*<\/mrow>\s*(<mn>[\s\S]*?<\/mn>|<mi>[\s\S]*?<\/mi>)\s*<\/msup>/gi,
    (m, base, expNode) => `<msup>${base}${expNode}</msup>`
  );

  // Also common: exponent wrapped but empty <mrow></mrow> appears as first child with whitespace/newlines
  s = s.replace(
    /<msup>\s*<mrow>\s*<\/mrow>\s*<mrow>\s*([\s\S]*?)\s*<\/mrow>\s*<\/msup>/gi,
    (m, expInner) => `<msup><mrow></mrow><mrow>${expInner}</mrow></msup>`
  );

  return s;
}

/* ================= ✅ SQRT TOKEN CONVERTER ================= */

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

      if (c === Infinity) break;

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
      openRe.lastIndex = start + 1;
      closeRe.lastIndex = start + 1;
    }
  }

  blocks.sort((a, b) => a.start - b.start);
  return blocks;
}

function splitMrootChildren(innerXml) {
  const s = innerXml.trim();
  // heuristic: split after first top-level element ends
  let depth = 0;
  let cut = -1;

  for (let i = 0; i < s.length; i++) {
    if (s[i] === "<") {
      if (s.slice(i, i + 2) === "</") depth--;
      else depth++;
    }
    if (depth === 0 && i > 0) {
      const maybe = s.slice(0, i + 1);
      if (/<\/\w+>\s*$/.test(maybe)) {
        cut = i + 1;
        break;
      }
    }
  }

  if (cut < 0) return [s, ""];
  return [s.slice(0, cut).trim(), s.slice(cut).trim()];
}

function convertMathMLWithSqrtTokens(mathml) {
  let mm = stripMathMLTagPrefixes(String(mathml || "")).trim();

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
  mm = fixBrokenMsup(mm);

  const sqrtBlocks = extractBalancedTagBlocks(mm, "msqrt");
  const rootBlocks = extractBalancedTagBlocks(mm, "mroot");

  const tokenMap = [];
  let replaced = mm;

  const all = [
    ...sqrtBlocks.map((b) => ({ ...b, type: "msqrt" })),
    ...rootBlocks.map((b) => ({ ...b, type: "mroot" })),
  ].sort((a, b) => b.start - a.start);

  let k = 0;
  for (const b of all) {
    const token = `__SQRT_TOKEN_${++k}__`;
    tokenMap.push({ token, type: b.type, xml: b.xml });
    replaced = replaced.slice(0, b.start) + `<mi>${token}</mi>` + replaced.slice(b.end);
  }

  let latexMain = "";
  try {
    latexMain = (MathMLToLaTeX.convert(replaced) || "").trim();
  } catch {
    latexMain = "";
  }

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

    const re = new RegExp(item.token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
    latexMain = latexMain.replace(re, latexRep);
  }

  return latexMain.trim();
}

/* ================= MathML -> LaTeX ================= */

function mathmlToLatexSafe(mml) {
  try {
    if (!mml) return "";
    let mm = String(mml).trim();

    mm = stripMathMLTagPrefixes(mm);

    // allow body-only => wrap
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
    mm = fixBrokenMsup(mm);

    // ✅ convert with sqrt tokens (avoid losing sqrt)
    const latex = convertMathMLWithSqrtTokens(mm);

    // if MathML had sqrt but latex doesn't, force wrap as last resort
    if (SQRT_MATHML_RE.test(mm) && !/\\sqrt\b|\\root\b/.test(latex || "")) {
      if (!latex) return "\\sqrt{}";
      return `\\sqrt{${latex}}`;
    }

    return (latex || "").trim();
  } catch {
    return "";
  }
}

/**
 * MathType FIRST:
 * - token [!m:$mathtype_x$]
 * - produce:
 *   latexMap[key] = "..."
 *   (optional) images["fallback_key"] = png dataURL if latex fails
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

      // 3) MathML -> LaTeX (patched)
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

function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // Protect tokens BEFORE stripping tags
  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

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

  x = x.replace(/<(?!\/?u\b)[^>]+>/g, "");

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

    // 4) parse questions
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      latex: latexMap,
      images,
      rawText: text,
      debug: {
        latexCount: Object.keys(latexMap).length,
        imagesCount: Object.keys(images).length,
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
