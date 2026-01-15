import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import { execFile, execFileSync } from "child_process";

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

/**
 * Convert EMF/WMF -> PNG using Inkscape
 */
function inkscapeConvertToPng(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    execFile(
      "inkscape",
      [inputPath, "--export-type=png", `--export-filename=${outputPath}`],
      { timeout: 20000 },
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
    const png = fs.readFileSync(outPath);
    return png;
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {}
  }
}

/**
 * Try extract embedded MathML from MathType OLE binary (scan <math>...</math>)
 */
function extractMathMLFromOle(buf) {
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
 * MathType FIRST:
 * - Replace <w:object> with [!m:$mathtype_x$]
 * - Extract MathML from oleObject.bin
 * - If no MathML: fallback preview rid and return fallback image dataURL in images[fallback_key]
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

  const mathMap = {};

  await Promise.all(
    Object.entries(found).map(async ([key, info]) => {
      const oleFull = normalizeTargetToWordPath(info.oleTarget);
      const oleBuf = await getZipEntryBuffer(zipFiles, oleFull);

      if (oleBuf) {
        const mml = extractMathMLFromOle(oleBuf);
        if (mml && mml.trim()) {
          mathMap[key] = mml;
          return;
        }
      }

      // fallback preview image
      if (info.previewRid) {
        const t = rels.get(info.previewRid);
        if (t) {
          const imgFull = normalizeTargetToWordPath(t);
          const imgBuf = await getZipEntryBuffer(zipFiles, imgFull);
          if (imgBuf) {
            const mime = guessMimeFromFilename(imgFull);

            // convert emf/wmf -> png
            if (mime === "image/emf" || mime === "image/wmf") {
              try {
                const pngBuf = await maybeConvertEmfWmfToPng(imgBuf, imgFull);
                if (pngBuf) {
                  images[`fallback_${key}`] =
                    `data:image/png;base64,${pngBuf.toString("base64")}`;
                  mathMap[key] = ""; // no MathML
                  return;
                }
              } catch {}
            }

            images[`fallback_${key}`] =
              `data:${mime};base64,${imgBuf.toString("base64")}`;
          }
        }
      }

      mathMap[key] = "";
    })
  );

  return { outXml: docXml, mathMap };
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

/**
 * ✅ KEY FIX: Convert Word XML to text BUT KEEP underline markers
 *
 * - Keeps tokens [!m:$..$] and [!img:$..$]
 * - Preserves underline run by wrapping run text with [[U]]...[[/U]]
 *
 * NOTE: underline in Word = <w:rPr>...<w:u .../>...</w:rPr>
 */
function wordXmlToTextKeepTokensAndUnderline(docXml) {
  // Normalize paragraph + breaks + tabs
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // Protect our tokens before we mess with tags
  x = x.replace(/\[!m:\$(.*?)\$\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$(.*?)\$\]/g, "___IMG_TOKEN___$1___END___");

  // Convert runs: <w:r> ... </w:r>
  // If run has underline (<w:u .../>) => wrap its text with [[U]]..[[/U]]
  x = x.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
    const isUnderline =
      /<w:rPr[\s\S]*?<w:u\b[^>]*\/>[\s\S]*?<\/w:rPr>/.test(run) ||
      /<w:u\b[^>]*\/>/.test(run);

    // Extract all <w:t> segments in this run (can be multiple)
    let txt = "";
    run.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => {
      txt += inner;
      return m;
    });

    // If run contains no text, ignore it
    if (!txt) return "";

    // Wrap underline marker
    if (isUnderline) {
      return `[[U]]${txt}[[/U]]`;
    }
    return txt;
  });

  // Remove leftover tags
  x = x.replace(/<[^>]+>/g, "");

  // Restore our tokens
  x = x
    .replace(/___MATH_TOKEN___(.*?)___END___/g, "[!m:$$$1$$]")
    .replace(/___IMG_TOKEN___(.*?)___END___/g, "[!img:$$$1$$]");

  // Decode entities + cleanup
  x = decodeXmlEntities(x)
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return x;
}

/**
 * ✅ Parse questions:
 * - MCQ A-D
 * - Detect correct answer from underline marker [[U]] in choice text
 */
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
      no: null,
    };

    const mNo = block.match(/^Câu\s*(\d+)\./);
    q.no = mNo ? parseInt(mNo[1], 10) : null;

    const [main, solution] = block.split(/Lời giải/i);
    q.solution = solution ? solution.trim() : "";

    // Parse choices A-D robust
    const choiceRe =
      /(\*?)([A-D])\.\s([\s\S]*?)(?=\n\*?[A-D]\.\s|\nLời giải|$)/g;
    let m;
    while ((m = choiceRe.exec(main))) {
      const starred = m[1] === "*";
      const label = m[2];
      const content = (m[3] || "").trim();

      if (starred) q.correct = label;

      // ✅ underline-based correct
      if (!q.correct && content.includes("[[U]]")) {
        q.correct = label;
      }

      q.choices.push({ label, text: content });
    }

    // content = stem
    const splitAtA = main.split(/\n\*?A\.\s/);
    q.content = (splitAtA[0] || "").trim();

    // fallback: "Chọn C" in solution
    if (!q.correct && q.solution) {
      const pick = q.solution.match(/\bChọn\s*([A-D])\b/i);
      if (pick) q.correct = pick[1].toUpperCase();
    }

    questions.push(q);
  }

  return questions;
}

/* ================= API ================= */

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
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

    // ✅ MathType first (fallback preview + convert emf/wmf -> png)
    const images = {};
    const mathTok = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mathTok.outXml;
    const mathMap = mathTok.mathMap;

    // ✅ Then normal images (also convert emf/wmf -> png)
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // ✅ Convert XML -> text but keep underline markers
    const text = wordXmlToTextKeepTokensAndUnderline(docXml);

    // ✅ Parse quiz
    const questions = parseQuestions(text);

    // NOTE: bạn đang render LaTeX ở HTML từ data.latex.
    // Nếu hiện tại bạn chưa có converter MathML->LaTeX thì tạm trả latexMap rỗng,
    // HTML sẽ fallback ảnh công thức (fallback_mathtype_x).
    // Khi bạn đã có latex converter, chỉ cần set latexMap[key] = "...".
    const latexMap = {}; // <-- bạn có thể nối vào pipeline LaTeX sau

    res.json({
      ok: true,
      total: questions.length,
      questions,
      // trả cả math (MathML) để debug nếu cần
      math: mathMap,
      latex: latexMap,
      images,
      rawText: text,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

app.get("/debug-inkscape", (_, res) => {
  try {
    const v = execFileSync("inkscape", ["--version"]).toString();
    res.type("text/plain").send(v);
  } catch (e) {
    res.status(500).type("text/plain").send("NO INKSCAPE");
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
