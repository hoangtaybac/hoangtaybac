// server.js
import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import { execFile, execFileSync } from "child_process";
import { MathMLToLaTeX } from "mathml-to-latex";
import { createRequire } from "module";

const require = createRequire(import.meta.url);

const app = express();
app.use(cors());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/* ================= OMML -> MathML (optional) =================
 - Uses OMML2MML.XSL (Microsoft/Office XSLT) placed next to this file.
 - Requires native libs: libxslt, libxmljs (npm packages).
 - If missing, OMML conversion is skipped silently.
 - Precompiles XSLT, caches conversions, limits concurrency.
================================================================*/

const OMML_XSL_PATH = path.join(__dirname, "OMML2MML.XSL");
let ommlAvailable = false;
let ommlStylesheet = null;
let libxslt = null;
let libxmljs = null;
const ommlCache = new Map(); // fragment -> latex

try {
  if (fs.existsSync(OMML_XSL_PATH)) {
    // try to require native libs
    try {
      libxslt = require("libxslt");
      libxmljs = require("libxmljs");
      const xslStr = fs.readFileSync(OMML_XSL_PATH, "utf8");
      try {
        // try parse both ways for compatibility
        ommlStylesheet = libxslt.parse(xslStr);
      } catch (e) {
        ommlStylesheet = libxslt.parse(libxmljs.parseXml(xslStr));
      }
      ommlAvailable = !!ommlStylesheet;
      if (ommlAvailable) console.log("✅ OMML2MML stylesheet loaded. OMML conversion enabled.");
    } catch (e) {
      ommlAvailable = false;
      console.warn("⚠️ OMML2MML.XSL is present but libxslt/libxmljs are not installed. OMML conversion disabled.");
    }
  } else {
    ommlAvailable = false;
    console.log("ℹ️ OMML2MML.XSL not found. OMML conversion disabled.");
  }
} catch (e) {
  ommlAvailable = false;
  console.warn("⚠️ Error checking OMML2MML.XSL:", e && e.message);
}

function convertOmmlFragmentToMathML(ommlFragment) {
  // wrap fragment in minimal document expected by stylesheet
  const wrapper = `<?xml version="1.0" encoding="UTF-8"?>
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    ${ommlFragment}
  </w:document>`;
  // parse & apply
  const doc = libxmljs.parseXml(wrapper);
  const result = ommlStylesheet.apply(doc);
  return String(result || "");
}

// small async pool to limit concurrency
async function asyncPool(limit, array, iteratorFn) {
  const ret = [];
  const executing = [];
  for (const item of array) {
    const p = Promise.resolve().then(() => iteratorFn(item));
    ret.push(p);

    if (limit <= 0) {
      // run unlimited
      continue;
    }

    const e = p.then(() => {
      executing.splice(executing.indexOf(e), 1);
    });
    executing.push(e);
    if (executing.length >= limit) {
      await Promise.race(executing);
    }
  }
  return Promise.all(ret);
}

/**
 * Find OMML fragments (<m:oMath> / <m:oMathPara>) and convert to MathML -> LaTeX.
 * Replace fragment occurrences with token [!m:$$omml_x$$] and fill latexMap.
 *
 * Returns { outXml, ommlLatexMap }
 */
async function tokenizeOmmlFirst(docXml, latexMap) {
  if (!ommlAvailable) {
    return { outXml: docXml, ommlLatexMap: {} };
  }

  const OMML_RE = /<m:oMathPara[\s\S]*?<\/m:oMathPara>|<m:oMath[\s\S]*?<\/m:oMath>/g;
  const found = [];
  let match;
  let idx = 0;

  // Collect fragments and their positions
  while ((match = OMML_RE.exec(docXml)) !== null) {
    const frag = match[0];
    const key = `omml_${++idx}`;
    found.push({ key, frag });
  }

  if (found.length === 0) return { outXml: docXml, ommlLatexMap: {} };

  // process fragments with concurrency (avoid blocking)
  const concurrency = Math.max(1, os.cpus().length - 1);
  await asyncPool(concurrency, found, async (entry) => {
    try {
      const fragTrim = entry.frag.trim();
      if (ommlCache.has(fragTrim)) {
        latexMap[entry.key] = ommlCache.get(fragTrim);
        return;
      }
      // convert OMML -> MathML using stylesheet
      let mml = "";
      try {
        mml = convertOmmlFragmentToMathML(entry.frag);
      } catch (e) {
        console.warn(`[OMML] XSLT apply failed for ${entry.key}:`, e && e.message);
        mml = "";
      }
      // Convert MathML -> LaTeX (reuse existing function)
      const latex = mml ? mathmlToLatexSafe(mml) : "";
      latexMap[entry.key] = latex || "";
      ommlCache.set(fragTrim, latexMap[entry.key]);
    } catch (e) {
      console.error(`[OMML] unexpected error for ${entry.key}:`, e && e.message);
      latexMap[entry.key] = "";
    }
  });

  // Replace fragments in docXml with tokens (do a single replace per fragment key)
  // Note: some identical fragments may appear; we replace strictly first N matches sequentially
  let outXml = docXml;
  let replaceIdx = 0;
  outXml = outXml.replace(OMML_RE, () => {
    replaceIdx++;
    const key = `omml_${replaceIdx}`;
    return `[!m:$${key}$]`;
  });

  // Build omml map to return (subset of latexMap)
  const ommlLatexMap = {};
  for (let i = 1; i <= idx; i++) {
    const k = `omml_${i}`;
    if (Object.prototype.hasOwnProperty.call(latexMap, k)) {
      ommlLatexMap[k] = latexMap[k];
    }
  }

  return { outXml, ommlLatexMap };
}

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
  const candidates = [];
  try {
    candidates.push(buf.toString("utf8"));
  } catch {}
  try {
    candidates.push(buf.toString("utf16le"));
  } catch {}
  try {
    // thử cả utf16be (nhiều OLE dùng BE)
    const be = Buffer.from(buf);
    for (let i = 0; i + 1 < be.length; i += 2) {
      const a = be[i];
      be[i] = be[i + 1];
      be[i + 1] = a;
    }
    candidates.push(be.toString("utf16le"));
  } catch {}

  for (const txt of candidates) {
    if (!txt) continue;
    const i = txt.indexOf("<math");
    if (i !== -1) {
      const j = txt.indexOf("</math>", i);
      if (j !== -1) {
        const slice = txt.slice(i, j + 7);
        return slice;
      }
    }
    const unescaped = txt.replace(/&lt;/g, "<").replace(/&gt;/g, ">");
    const iu = unescaped.indexOf("<math");
    if (iu !== -1) {
      const ju = unescaped.indexOf("</math>", iu);
      if (ju !== -1) return unescaped.slice(iu, ju + 7);
    }
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

    const scriptPath = path.join(__dirname, "mt2mml.rb");

    const opts = {
      timeout: 120000,
      maxBuffer: 50 * 1024 * 1024,
      env: Object.assign({}, process.env, { LANG: "en_US.UTF-8" }),
      cwd: path.dirname(scriptPath),
    };

    execFile("ruby", [scriptPath, inPath], opts, (err, stdout, stderr) => {
      try {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      } catch {}
      if (err) {
        console.error("[rubyOleToMathML] ruby exec error:", err && err.code, err && err.message);
        if (stderr) console.error("[rubyOleToMathML] stderr:", String(stderr).slice(0, 4000));
        if (stdout) console.error("[rubyOleToMathML] stdout:", String(stdout).slice(0, 4000));
        return reject(new Error(stderr ? String(stderr).trim() : (err && err.message) || "ruby failed"));
      }
      const out = String(stdout || "").trim();
      if (!out) console.warn("[rubyOleToMathML] ruby returned empty stdout");
      return resolve(out);
    });
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
 * ✅ XỬ LÝ CĂN (sqrt) “cứng”:
 */
function fixSqrtLatex(latex, mathmlMaybe = "") {
  let s = String(latex || "");

  s = s.replace(/√\s*\(\s*([\s\S]*?)\s*\)/g, "\\sqrt{$1}");
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

/* ================= MathType OLE processing ================= */

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
      try {
        const oleFull = normalizeTargetToWordPath(info.oleTarget);
        console.log(`[tokenizeMathType] key=${key} oleTarget=${oleFull} previewRid=${info.previewRid}`);
        const oleBuf = await getZipEntryBuffer(zipFiles, oleFull);
        if (!oleBuf) {
          console.warn(`[tokenizeMathType] no ole buffer for key=${key}`);
        } else {
          console.log(`[tokenizeMathType] oleBuf.length=${oleBuf.length} for key=${key}`);
        }

        // 1) try scan MathML inside OLE
        let mml = "";
        if (oleBuf) {
          mml = extractMathMLFromOleScan(oleBuf) || "";
          if (mml) console.log(`[tokenizeMathType] extracted MathML (scan) len=${mml.length} for key=${key}`);
        }

        // 2) fallback ruby convert (MTEF inside OLE)
        if (!mml && oleBuf) {
          try {
            console.log(`[tokenizeMathType] calling ruby mt2mml for key=${key}`);
            mml = await rubyOleToMathML(oleBuf);
            if (mml) console.log(`[tokenizeMathType] ruby produced MathML len=${mml.length} for key=${key}`);
          } catch (e) {
            console.warn(`[tokenizeMathType] ruby conversion failed for key=${key}: ${e && e.message}`);
            mml = "";
          }
        }

        // 3) MathML -> LaTeX (✅ có postprocess căn/cases)
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
                    images[`fallback_${key}`] = `data:image/png;base64,${pngBuf.toString(
                      "base64"
                    )}`;
                    latexMap[key] = "";
                    console.log(`[tokenizeMathType] fallback image converted emf/wmf -> png for key=${key}`);
                    return;
                  }
                } catch (e) {
                  console.warn(`[tokenizeMathType] inkscape convert error for key=${key}: ${e && e.message}`);
                }
              }
              images[`fallback_${key}`] = `data:${mime};base64,${imgBuf.toString("base64")}`;
              console.log(`[tokenizeMathType] using preview image for key=${key} mime=${mime}`);
            }
          }
        }

        latexMap[key] = "";
      } catch (err) {
        console.error(`[tokenizeMathType] unexpected error for key=${key}:`, err && err.message);
        latexMap[key] = "";
      }
    })
  );

  return { outXml: docXml, latexMap };
}

/* ================= Tokenize normal images AFTER MathType ================= */

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

/* ================== EXAM PARSER (unchanged) ================== */

// ... (remaining parsing functions unchanged) ...
// For brevity in this snippet we assume the rest of your original parsing
// functions (stripTagsToPlain, detectHasMCQ, detectHasTF4, ...,
// parseExamFromText, legacyQuestionsFromExam) are unchanged and follow here.
// (In your actual file they are kept as in your original code.)

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

    // NEW STEP: OMML -> MathML -> LaTeX (if available) (minimal & optional)
    const images = {};
    const initialLatexMap = {};
    if (ommlAvailable) {
      try {
        const ommlRes = await tokenizeOmmlFirst(docXml, initialLatexMap);
        docXml = ommlRes.outXml;
        // initialLatexMap keys are already in initialLatexMap; they'll be merged below
      } catch (e) {
        console.warn("OMML conversion failed (continuing):", e && e.message);
      }
    }

    // 1) MathType -> LaTeX (and fallback images) ✅ giữ pipeline B
    const mt = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mt.outXml;
    const latexMap = Object.assign({}, initialLatexMap, mt.latexMap);

    // 2) normal images ✅ giữ pipeline B
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // 3) text ✅ giữ token + underline
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) NEW: parse exam output (mcq/tf4/short)
    const exam = parseExamFromText(text);

    // 5) legacy questions output (mcq only) for backward compatibility
    const questions = legacyQuestionsFromExam(exam);

    res.json({
      ok: true,
      total: exam.questions.length,
      exam,
      // giữ field cũ
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
