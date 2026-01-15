import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import crypto from "crypto";
import fs from "fs";
import os from "os";
import path from "path";
import { execFileSync } from "child_process";

const app = express();
app.use(cors());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/**
 * In-memory store (MVP).
 * - images: filename -> { buf, ext }
 * - wmfPngCache: filename -> pngBuf
 */
const MEM = {
  images: new Map(),
  wmfPngCache: new Map(),
};

/* ===================== Helpers ===================== */

function stripTags(xml) {
  return xml.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

function parseRels(relsXml) {
  const map = new Map();
  const re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  let m;
  while ((m = re.exec(relsXml))) {
    map.set(m[1], m[2]);
  }
  return map;
}

function extOf(name = "") {
  const m = name.toLowerCase().match(/\.([a-z0-9]+)$/);
  return m ? m[1] : "";
}

function contentTypeByExt(ext) {
  switch (ext) {
    case "png": return "image/png";
    case "jpg":
    case "jpeg": return "image/jpeg";
    case "gif": return "image/gif";
    case "svg": return "image/svg+xml";
    case "webp": return "image/webp";
    case "bmp": return "image/bmp";
    default: return "application/octet-stream";
  }
}

/**
 * Try extract embedded MathML from MathType OLE binary.
 * Many MathType OLE objects embed a MathML translation.
 */
function extractMathMLFromOle(buf) {
  // UTF-8 scan
  const u8 = buf.toString("utf8");
  let i = u8.indexOf("<math");
  if (i !== -1) {
    const j = u8.indexOf("</math>", i);
    if (j !== -1) return u8.slice(i, j + 7);
  }

  // UTF-16LE scan (common in OLE streams)
  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    const j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }

  return null;
}

function wmfToPngBuffer(wmfBuf) {
  // Cache by hash to avoid repeated convert cost
  const h = crypto.createHash("md5").update(wmfBuf).digest("hex");
  if (MEM.wmfPngCache.has(h)) return MEM.wmfPngCache.get(h);

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "wmf-"));
  const inPath = path.join(tmpDir, "in.wmf");
  const outPath = path.join(tmpDir, "out.png");
  fs.writeFileSync(inPath, wmfBuf);

  // ImageMagick convert (provided by nixpacks.toml)
  execFileSync("convert", [inPath, outPath], { stdio: "ignore" });

  const png = fs.readFileSync(outPath);
  fs.rmSync(tmpDir, { recursive: true, force: true });

  MEM.wmfPngCache.set(h, png);
  return png;
}

/**
 * Tokenize OLE MathType blocks:
 * - If MathML extracted => [!m:$mathtype_x$] and mathMap[key]=MathML
 * - Else fallback to preview image => [img:$eq_x$] and imagesMap[eq]=url
 */
async function tokenizeMathTypeOle(docXml, rels, zipFiles, baseUrl) {
  let mathIdx = 0;
  let eqIdx = 0;

  const mathMap = {};
  const imagesMap = {};

  const getZipBuffer = async (p) => {
    const f = zipFiles.find((x) => x.path === p);
    return f ? await f.buffer() : null;
  };

  // Replace <w:object>...</w:object>
  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;
  const jobs = [];

  const outXml = docXml.replace(OBJECT_RE, (block) => {
    // OLE object relationship (embeddings/oleObjectX.bin)
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid); // e.g. "embeddings/oleObject1.bin"
    if (!oleTarget) return block;

    // preview image relationship (media/imageX.wmf/png/..)
    const img = block.match(/<v:imagedata\b[^>]*\br:id="([^"]+)"/);
    const imgRid = img ? img[1] : null;
    const imgTarget = imgRid ? rels.get(imgRid) : null; // e.g. "media/image1.wmf"

    // Create placeholder key now, decide later in job
    const tempKey = `__TMP__${crypto.randomUUID?.() || Math.random().toString(16).slice(2)}`;

    jobs.push(
      (async () => {
        const olePath = oleTarget.startsWith("word/") ? oleTarget : `word/${oleTarget}`;
        const oleBuf = await getZipBuffer(olePath);
        const mml = oleBuf ? extractMathMLFromOle(oleBuf) : null;

        if (mml) {
          const key = `mathtype_${++mathIdx}`;
          mathMap[key] = mml;
          // Replace tempKey in docXml later
          mathMap[tempKey] = { kind: "math", key };
          return;
        }

        // Fallback to image token
        if (imgTarget && imgTarget.startsWith("media/")) {
          const filename = imgTarget.replace("media/", "");
          const key = `eq_${++eqIdx}`;
          imagesMap[key] = `${baseUrl}/img/${encodeURIComponent(filename)}`;
          mathMap[tempKey] = { kind: "img", key };
        } else {
          // No image relationship -> empty
          mathMap[tempKey] = { kind: "empty", key: "" };
        }
      })()
    );

    // Put temporary marker; will be replaced after jobs resolve
    return `[__${tempKey}__]`;
  });

  await Promise.all(jobs);

  // Replace temp markers with actual tokens
  let finalXml = outXml;
  for (const [k, v] of Object.entries(mathMap)) {
    if (!k.startsWith("__TMP__")) continue;
    const marker = `[__${k}__]`;
    if (v.kind === "math") {
      finalXml = finalXml.split(marker).join(`[!m:$${v.key}$]`);
    } else if (v.kind === "img") {
      finalXml = finalXml.split(marker).join(`[img:$${v.key}$]`);
    } else {
      finalXml = finalXml.split(marker).join("");
    }
    delete mathMap[k]; // remove temp
  }

  return { xml: finalXml, mathMap, imagesMap };
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

    main.replace(/\*?([A-D])\.\s([^A-D]*)/g, (m, label, content) => {
      if (m.startsWith("*")) q.correct = label;
      q.choices.push({ label, text: content.trim() });
      return m;
    });

    q.content = main.split(/A\./)[0].trim();
    questions.push(q);
  }

  return questions;
}

/* ===================== API ===================== */

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    MEM.images.clear();
    MEM.wmfPngCache.clear();

    const zip = await unzipper.Open.buffer(req.file.buffer);

    const docEntry = zip.files.find((f) => f.path === "word/document.xml");
    const relEntry = zip.files.find((f) => f.path === "word/_rels/document.xml.rels");
    if (!docEntry || !relEntry) throw new Error("Missing document.xml or document.xml.rels");

    let docXml = (await docEntry.buffer()).toString("utf8");
    const relsXml = (await relEntry.buffer()).toString("utf8");
    const rels = parseRels(relsXml);

    // Load media into memory (for image fallback)
    for (const f of zip.files) {
      if (f.path.startsWith("word/media/")) {
        const filename = f.path.replace("word/media/", "");
        const buf = await f.buffer();
        MEM.images.set(filename, { buf, ext: extOf(filename) });
      }
    }

    const baseUrl = `${req.protocol}://${req.get("host")}`;

    // Tokenize OLE MathType => math token (preferred) or image token fallback
    const { xml, mathMap, imagesMap } = await tokenizeMathTypeOle(docXml, rels, zip.files, baseUrl);
    docXml = xml;

    // Strip to text (keep tokens)
    const text = stripTags(docXml);

    // Parse questions
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      math: mathMap,     // mathtype_x -> <math>...</math>
      images: imagesMap, // eq_x -> URL (/img/filename)
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Serve media images; WMF auto-convert to PNG for browser
app.get("/img/:name", (req, res) => {
  const name = req.params.name;
  const item = MEM.images.get(name);
  if (!item) return res.status(404).send("not found");

  const ext = item.ext;

  try {
    if (ext === "wmf") {
      const png = wmfToPngBuffer(item.buf);
      res.setHeader("Content-Type", "image/png");
      res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
      return res.send(png);
    }

    res.setHeader("Content-Type", contentTypeByExt(ext));
    res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
    return res.send(item.buf);
  } catch (e) {
    console.error("img serve error:", e);
    return res.status(500).send("image convert error");
  }
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
