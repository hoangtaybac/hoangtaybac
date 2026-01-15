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
app.use(express.json({ limit: "25mb" }));

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

function mathmlToLatexSafe(mml) {
  try {
    if (!mml || !mml.includes("<math")) return "";
    const latex = MathMLToLaTeX.convert(mml);
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

      // 3) MathML -> LaTeX
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
 * Convert Word document.xml -> text
 * - giữ token [!m:$$key$$], [!img:$$key$$]
 * - GIỮ underline bằng cách bọc <u>...</u>
 */
function wordXmlToTextKeepTokens(docXml) {
  // chuẩn hoá tab, br
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n");

  // giữ tokens trước khi parse run
  x = x.replace(/\[!m:\$(.*?)\$\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$(.*?)\$\]/g, "___IMG_TOKEN___$1___END___");

  // tách paragraph
  const paras = x.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const out = [];

  for (const p of paras) {
    const runs = p.match(/<w:r\b[\s\S]*?<\/w:r>/g) || [];
    let line = "";

    for (const r of runs) {
      // underline?
      let isUnderline = false;
      const uTag = r.match(/<w:u\b[^>]*\/>/);
      if (uTag) {
        const val = uTag[0].match(/w:val="([^"]+)"/)?.[1];
        isUnderline = !val || val.toLowerCase() !== "none";
      }

      // lấy text trong run
      const texts = [];
      const tRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
      let tm;
      while ((tm = tRe.exec(r))) texts.push(tm[1] ?? "");
      if (!texts.length) continue;

      let runText = texts.join("");

      // KHÔNG để underline "ăn" vào token math/img
      const parts = runText.split(
        /(___MATH_TOKEN___.*?___END___|___IMG_TOKEN___.*?___END___)/g
      );

      const rebuilt = parts
        .map((part) => {
          if (!part) return "";
          const isToken =
            part.startsWith("___MATH_TOKEN___") ||
            part.startsWith("___IMG_TOKEN___");
          if (isToken) return part;
          return isUnderline ? `<u>${part}</u>` : part;
        })
        .join("");

      line += rebuilt;
    }

    // decode entities
    line = decodeXmlEntities(line).replace(/\r/g, "");

    // chỉ cho phép <u> </u>, xoá tag lạ (nếu còn sót)
    line = line.replace(/<[^>]+>/g, (m) => {
      if (m === "<u>" || m === "</u>") return m;
      return "";
    });

    // bỏ dòng rỗng
    if (line.trim().length > 0) out.push(line);
  }

  let text = out.join("\n");
  text = text.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();

  // restore token về dạng bạn dùng
  text = text
    .replace(/___MATH_TOKEN___(.*?)___END___/g, "[!m:$$$1$$]")
    .replace(/___IMG_TOKEN___(.*?)___END___/g, "[!img:$$$1$$]");

  return text;
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

    // đáp án A-D có thể có underline dạng <u>...</u> trong text
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
      throw new Error("Missing word/document.xml or word/_rels/document.xml.rels");

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

    // 3) text giữ underline
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) parse câu hỏi
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      latex: latexMap,
      images,
      rawText: text,
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
