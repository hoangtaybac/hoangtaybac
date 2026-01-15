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

const DEBUG_MT = !!process.env.DEBUG_MT;

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

/* ================= Inkscape Convert EMF/WMF -> PNG (unchanged) ================= */

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
 * scan MathML embedded directly in OLE quickly
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
 * (kept as-is - requires mt2mml.rb in PATH)
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

/* ------------------ New helpers to handle OMML radical variants ------------------ */

/**
 * Strip common OMML wrapper tags like <m:r><m:t>... to their text content
 * Conservative: only remove wrappers expected from OMML
 */
function stripOmmlWrappers(s) {
  if (!s) return s;
  return s
    .replace(/<m:r\b[^>]*>/gi, "")
    .replace(/<\/m:r>/gi, "")
    .replace(/<m:t\b[^>]*>/gi, "")
    .replace(/<\/m:t>/gi, "")
    .replace(/<r\b[^>]*>/gi, "")
    .replace(/<\/r>/gi, "")
    .replace(/<t\b[^>]*>/gi, "")
    .replace(/<\/t>/gi, "");
}

/**
 * Heuristic normalization: convert OMML-style radical nodes to MathML equivalents.
 * - handle <m:rad> / <rad> variants (with nested <m:e>, <m:deg>, or wrappers)
 * - produce <msqrt> or <mroot> to help MathML->LaTeX convertor
 */
function normalizeRadicalsInMathML(mml) {
  if (!mml || typeof mml !== "string") return mml;
  let s = mml;

  // normalize newlines to spaces for our regex-based transforms
  s = s.replace(/\r?\n/g, " ");

  // PASS 1: <m:rad> ... </m:rad>
  s = s.replace(/<m:rad\b[^>]*>([\s\S]*?)<\/m:rad>/gi, (full, inner) => {
    const degMatch = inner.match(/<m:deg\b[^>]*>([\s\S]*?)<\/m:deg>/i);
    const eMatch = inner.match(/<m:e\b[^>]*>([\s\S]*?)<\/m:e>/i);
    const deg = degMatch ? stripOmmlWrappers(degMatch[1]).trim() : "";
    const e = eMatch ? stripOmmlWrappers(eMatch[1]).trim() : "";
    if (e && deg) return `<mroot>${e}${deg}</mroot>`;
    if (e) return `<msqrt>${e}</msqrt>`;
    const fallback = stripOmmlWrappers(inner).trim();
    if (fallback) return `<msqrt>${fallback}</msqrt>`;
    return full;
  });

  // PASS 2: <rad> ... </rad> (no prefix)
  s = s.replace(/<rad\b[^>]*>([\s\S]*?)<\/rad>/gi, (full, inner) => {
    const degMatch = inner.match(/<deg\b[^>]*>([\s\S]*?)<\/deg>/i);
    const eMatch = inner.match(/<e\b[^>]*>([\s\S]*?)<\/e>/i);
    const deg = degMatch ? stripOmmlWrappers(degMatch[1]).trim() : "";
    const e = eMatch ? stripOmmlWrappers(eMatch[1]).trim() : "";
    if (e && deg) return `<mroot>${e}${deg}</mroot>`;
    if (e) return `<msqrt>${e}</msqrt>`;
    const fallback = stripOmmlWrappers(inner).trim();
    if (fallback) return `<msqrt>${fallback}</msqrt>`;
    return full;
  });

  // PASS 3: nested/other <m:rad> patterns not caught earlier
  s = s.replace(/<m:rad\b[^>]*>([\s\S]*?)<\/m:rad>/gi, (full, inner) => {
    const degMatch = inner.match(/<m:deg\b[^>]*>([\s\S]*?)<\/m:deg>/i);
    const eMatch = inner.match(/<m:e\b[^>]*>([\s\S]*?)<\/m:e>/i);
    const deg = degMatch ? stripOmmlWrappers(degMatch[1]).trim() : "";
    const e = eMatch ? stripOmmlWrappers(eMatch[1]).trim() : "";
    if (e && deg) return `<mroot>${e}${deg}</mroot>`;
    if (e) return `<msqrt>${e}</msqrt>`;
    return full;
  });

  // PASS 4: remove m: prefix for common MathML tags to help converter
  s = s
    .replace(/<m:msqrt/gi, "<msqrt")
    .replace(/<\/m:msqrt/gi, "</msqrt")
    .replace(/<m:mroot/gi, "<mroot")
    .replace(/<\/m:mroot/gi, "</mroot")
    .replace(/<m:math/gi, "<math")
    .replace(/<\/m:math/gi, "</math")
    .replace(/<m:mn/gi, "<mn")
    .replace(/<\/m:mn/gi, "</mn")
    .replace(/<m:mi/gi, "<mi")
    .replace(/<\/m:mi/gi, "</mi")
    .replace(/<m:mo/gi, "<mo")
    .replace(/<\/m:mo/gi, "</mo")
    .replace(/<m:mrow/gi, "<mrow")
    .replace(/<\/m:mrow/gi, "</mrow");

  // PASS 5: collapse obvious wrapper runs into their inner text
  s = s.replace(/<m:r\b[^>]*>([\s\S]*?)<\/m:r>/gi, (_, inside) => stripOmmlWrappers(inside));
  s = s.replace(/<r\b[^>]*>([\s\S]*?)<\/r>/gi, (_, inside) => stripOmmlWrappers(inside));

  return s;
}

/**
 * Try MathML -> LaTeX with multiple fallback strategies.
 * Returns object: { latex: string, failedMml: string|null }
 */
function mathmlToLatexSafe(mml) {
  try {
    if (!mml || !mml.includes("<math")) return { latex: "", failedMml: null };

    // 1) try as-is
    try {
      const latex = MathMLToLaTeX.convert(mml);
      if (latex && String(latex).trim()) return { latex: (latex || "").trim(), failedMml: null };
    } catch {}

    // 2) normalize radicals/OMML wrappers
    const normalized = normalizeRadicalsInMathML(mml);
    if (normalized && normalized !== mml) {
      try {
        const latex2 = MathMLToLaTeX.convert(normalized);
        if (latex2 && String(latex2).trim()) return { latex: (latex2 || "").trim(), failedMml: null };
      } catch {}
    }

    // 3) more aggressive prefix removal + wrapper strip
    let step3 = normalized || mml;
    step3 = step3
      .replace(/<\/?m:([a-zA-Z0-9]+)/g, (m) => m.replace(/^<m:/, "<").replace(/^<\/m:/, "</"))
      .replace(/<m:r\b[^>]*>([\s\S]*?)<\/m:r>/gi, (_, inside) => stripOmmlWrappers(inside));
    try {
      const latex3 = MathMLToLaTeX.convert(step3);
      if (latex3 && String(latex3).trim()) return { latex: (latex3 || "").trim(), failedMml: null };
    } catch {}

    // 4) best-effort collapse to msqrt if radical-like found
    const hasRad = /<mroot|<msqrt|<rad|<m:rad/i.test(mml);
    if (hasRad) {
      const plain = stripOmmlWrappers(mml).replace(/<[^>]+>/g, "").trim();
      if (plain) {
        const guess = `<math><msqrt><mrow><mi>${plain}</mi></mrow></msqrt></math>`;
        try {
          const latex4 = MathMLToLaTeX.convert(guess);
          if (latex4 && String(latex4).trim()) return { latex: (latex4 || "").trim(), failedMml: null };
        } catch {}
      }
    }

    // all failed
    return { latex: "", failedMml: DEBUG_MT ? mml : null };
  } catch {
    return { latex: "", failedMml: DEBUG_MT ? mml : null };
  }
}

/* ================= Tokenize MathType objects (unchanged flow + fallback images) ================= */

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
  const failedMathML = {};

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

      // 3) MathML -> LaTeX via robust function
      if (mml) {
        const { latex, failedMml } = mathmlToLatexSafe(mml);
        if (latex) {
          latexMap[key] = latex;
          return;
        } else {
          if (failedMml) failedMathML[key] = failedMml;
        }
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

  return { outXml: docXml, latexMap, failedMathML };
}

/* ================= Tokenize normal images (unchanged) ================= */

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

/* ================= Text & Questions (unchanged) ================= */

function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // Protect tokens BEFORE stripping tags
  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

  // Convert each run <w:r> while preserving underline + token text
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

  // Restore tokens in a stable form (use $$ ... $$)
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
    const latexMap = mt.latexMap || {};
    const failedMathML = mt.failedMathML || {};

    // 2) normal images
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // 3) text: keep token + underline
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) parse questions
    const questions = parseQuestions(text);

    const resp = {
      ok: true,
      total: questions.length,
      questions,
      latex: latexMap,
      images,
      rawText: text,
    };

    if (DEBUG_MT) resp.failedMathML = failedMathML;

    res.json(resp);
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
