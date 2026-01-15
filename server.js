import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";

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

function guessMimeFromFilename(filename = "") {
  const ext = filename.split(".").pop()?.toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "bmp") return "image/bmp";
  if (ext === "webp") return "image/webp";
  if (ext === "svg") return "image/svg+xml";
  // emf/wmf: browser thường không render -> vẫn trả về để bạn thấy token, muốn giống Azota 100% cần convert sang png
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

function normalizeTargetToWordPath(target) {
  let t = (target || "").replace(/^(\.\.\/)+/, ""); // bỏ ../
  if (!t.startsWith("word/")) t = `word/${t}`;
  return t;
}

/**
 * Try extract embedded MathML from MathType OLE binary.
 */
function extractMathMLFromOle(buf) {
  // UTF-8 scan
  const utf8 = buf.toString("utf8");
  let i = utf8.indexOf("<math");
  if (i !== -1) {
    let j = utf8.indexOf("</math>", i);
    if (j !== -1) return utf8.slice(i, j + 7);
  }

  // UTF-16LE scan
  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    let j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }

  return null;
}

/**
 * 1) Tokenize MathType OLE FIRST
 * - Replace each <w:object>...</w:object> containing <o:OLEObject r:id="...">
 *   with token [!m:$mathtype_1$]
 * - Extract MathML from embeddings/oleObjectX.bin
 * - If no MathML => fallback to preview image RID inside the same block
 *   (supports both v:imagedata r:id and a:blip r:embed)
 * - Store fallback image as images["fallback_mathtype_1"] = dataURL
 */
async function tokenizeMathTypeOleFirst(docXml, rels, zipFiles, images) {
  let idx = 0;
  const mathMap = {};

  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;

  docXml = docXml.replace(OBJECT_RE, (block) => {
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid);
    if (!oleTarget) return block;

    // preview rid can be vml or drawing blip
    const vmlRid = block.match(
      /<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/
    );
    const blipRid = block.match(
      /<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/
    );
    const previewRid = vmlRid?.[1] || blipRid?.[1] || null;

    const key = `mathtype_${++idx}`;

    // We'll do the heavy extraction later by pushing jobs into a global list
    // but because replace callback is sync, we mark placeholders in a side table.
    mathMap[key] = { oleTarget, previewRid };

    return `[!m:$${key}$]`;
  });

  // Execute extraction for all keys
  const entries = Object.entries(mathMap);
  const finalMath = {};

  await Promise.all(
    entries.map(async ([key, info]) => {
      const oleFull = normalizeTargetToWordPath(info.oleTarget);
      const buf = await getZipEntryBuffer(zipFiles, oleFull);

      if (buf) {
        const mml = extractMathMLFromOle(buf);
        if (mml && mml.trim()) {
          finalMath[key] = mml;
          return;
        }
      }

      // fallback image
      if (info.previewRid) {
        const t = rels.get(info.previewRid);
        if (t) {
          const imgFull = normalizeTargetToWordPath(t);
          const imgbuf = await getZipEntryBuffer(zipFiles, imgFull);
          if (imgbuf) {
            const mime = guessMimeFromFilename(imgFull);
            images[`fallback_${key}`] = `data:${mime};base64,${imgbuf.toString(
              "base64"
            )}`;
          }
        }
      }

      finalMath[key] = ""; // no MathML
    })
  );

  return { outXml: docXml, mathMap: finalMath };
}

/**
 * 2) Tokenize normal images AFTER MathType
 * - drawing: <a:blip r:embed="rIdX"/>
 * - vml: <v:imagedata r:id="rIdY"/>
 * Replace by token [!img:$img_1$] and store images map.
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
        imgMap[key] = `data:${mime};base64,${buf.toString("base64")}`;
      })()
    );
  };

  // a:blip
  docXml = docXml.replace(
    /<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/g,
    (m, rid) => {
      const key = `img_${++idx}`;
      schedule(rid, key);
      return `[!img:$${key}$]`;
    }
  );

  // v:imagedata
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
 * 3) Convert Word XML to TEXT but keep tokens [!m:...] [!img:...]
 * - keep new line from </w:p>, <w:br/>
 * - keep tab from <w:tab/>
 */
function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // Protect tokens before stripping tags
  x = x.replace(/\[!m:\$(.*?)\$\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$(.*?)\$\]/g, "___IMG_TOKEN___$1___END___");

  // Replace <w:t>...</w:t> with inner text
  x = x.replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g, (seg) => {
    const mm = seg.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/);
    return mm ? mm[1] : "";
  });

  // Strip remaining tags
  x = x.replace(/<[^>]+>/g, "");

  // Restore tokens
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

    const choiceRe = /(\*?)([A-D])\.\s([\s\S]*?)(?=\n\*?[A-D]\.\s|\nLời giải|$)/g;
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

    // ✅ IMPORTANT: MathType FIRST so preview rid is still inside <w:object>
    const images = {};
    const mathTok = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mathTok.outXml;
    const mathMap = mathTok.mathMap;

    // Then tokenize remaining images
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // XML -> text keep tokens
    const text = wordXmlToTextKeepTokens(docXml);

    // Parse quiz blocks
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      math: mathMap,
      images,
      rawText: text, // debug nếu cần
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
