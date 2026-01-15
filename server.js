import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import path from "path";

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
  // docx đôi khi là emf/wmf (browser không hiển thị native). Bạn cần convert nếu gặp.
  if (ext === "emf") return "image/emf";
  if (ext === "wmf") return "image/wmf";
  return "application/octet-stream";
}

/**
 * Decode cơ bản entity trong XML (đủ cho docx text).
 */
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

/**
 * Try extract embedded MathML from MathType OLE binary.
 * MathType thường có MathML “dịch sẵn” trong OLE.
 */
function extractMathMLFromOle(buf) {
  // UTF-8 scan
  const utf8 = buf.toString("utf8");
  let i = utf8.indexOf("<math");
  if (i !== -1) {
    let j = utf8.indexOf("</math>", i);
    if (j !== -1) return utf8.slice(i, j + 7);
  }

  // UTF-16LE scan (rất hay gặp trong OLE)
  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    let j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }

  return null;
}

/**
 * Lấy buffer file trong zip theo path
 */
async function getZipEntryBuffer(zipFiles, p) {
  const f = zipFiles.find((x) => x.path === p);
  if (!f) return null;
  return await f.buffer();
}

/**
 * 1) Tokenize IMAGE:
 * - drawing: <a:blip r:embed="rIdX" .../>
 * - vml: <v:imagedata r:id="rIdY" .../>
 *
 * Thay chúng bằng token: [!img:$img_1$]
 * và tạo imgMap: img_1 -> dataURL
 */
async function tokenizeImages(docXml, rels, zipFiles) {
  let idx = 0;
  const imgMap = {};
  const jobs = [];

  function scheduleImageJob(rid, key) {
    const target = rels.get(rid);
    if (!target) return;

    // target thường kiểu "media/image1.png" hoặc "../media/..."
    // normalize về word/...
    let normalized = target.replace(/^(\.\.\/)+/, ""); // bỏ ../
    if (!normalized.startsWith("word/")) normalized = `word/${normalized}`;

    jobs.push(
      (async () => {
        const buf = await getZipEntryBuffer(zipFiles, normalized);
        if (!buf) return;

        const mime = guessMimeFromFilename(normalized);
        // Nếu là emf/wmf: browser khó render -> vẫn trả về để bạn biết, muốn chuẩn thì cần convert sang png.
        const b64 = buf.toString("base64");
        imgMap[key] = `data:${mime};base64,${b64}`;
      })()
    );
  }

  // a:blip r:embed
  docXml = docXml.replace(/<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/>/g, (m, rid) => {
    const key = `img_${++idx}`;
    scheduleImageJob(rid, key);
    return `[!img:$${key}$]`;
  });

  // v:imagedata r:id
  docXml = docXml.replace(/<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/g, (m, rid) => {
    const key = `img_${++idx}`;
    scheduleImageJob(rid, key);
    return `[!img:$${key}$]`;
  });

  await Promise.all(jobs);
  return { outXml: docXml, imgMap };
}

/**
 * 2) Tokenize MathType OLE:
 * - Tìm <w:object>...</w:object> có <o:OLEObject r:id="...">
 * - Thay toàn block bằng [!m:$mathtype_1$] (ưu tiên MathML)
 * - Nếu không extract được MathML => fallback ảnh preview nếu có v:imagedata trong block
 */
async function tokenizeMathTypeOle(docXml, rels, zipFiles, imgMap) {
  let idx = 0;
  const mathMap = {};
  const jobs = [];

  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;

  docXml = docXml.replace(OBJECT_RE, (block) => {
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid); // embeddings/oleObject1.bin
    if (!oleTarget) return block;

    // Nếu có preview image rid trong cùng block thì lấy để fallback
    const imgRidMatch = block.match(/<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/);
    const previewRid = imgRidMatch ? imgRidMatch[1] : null;

    const key = `mathtype_${++idx}`;

    jobs.push(
      (async () => {
        let full = oleTarget.replace(/^(\.\.\/)+/, "");
        if (!full.startsWith("word/")) full = `word/${full}`;

        const buf = await getZipEntryBuffer(zipFiles, full);
        if (!buf) {
          mathMap[key] = "";
          return;
        }

        const mml = extractMathMLFromOle(buf);
        if (mml) {
          mathMap[key] = mml;
          return;
        }

        // Fallback: nếu OLE không có MathML -> dùng ảnh preview (nếu có)
        if (previewRid) {
          const target = rels.get(previewRid);
          if (target) {
            let normalized = target.replace(/^(\.\.\/)+/, "");
            if (!normalized.startsWith("word/")) normalized = `word/${normalized}`;
            const imgbuf = await getZipEntryBuffer(zipFiles, normalized);
            if (imgbuf) {
              const mime = guessMimeFromFilename(normalized);
              imgMap[`fallback_${key}`] = `data:${mime};base64,${imgbuf.toString("base64")}`;
              mathMap[key] = ""; // báo client: không có MathML, dùng fallback ảnh
              return;
            }
          }
        }

        mathMap[key] = "";
      })()
    );

    return `[!m:$${key}$]`;
  });

  await Promise.all(jobs);
  return { outXml: docXml, mathMap };
}

/**
 * 3) Convert Word XML -> TEXT nhưng giữ token [!m:...] [!img:...]
 * - giữ xuống dòng theo </w:p>
 * - giữ tab theo <w:tab/>
 * - giữ line break theo <w:br/>
 * - lấy text từ <w:t>...</w:t>
 */
function wordXmlToTextKeepTokens(docXml) {
  // Đổi paragraph + break + tab thành ký hiệu text
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // Lấy nội dung trong w:t
  const texts = [];
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(x))) {
    texts.push(m[1]);
  }

  // Nhưng token [!m:$...$] / [!img:$...$] đang nằm ngoài w:t, nên phải giữ chúng:
  // -> Trick: trước khi remove tag, ta replace token vào marker riêng rồi strip.
  // Ở trên token vẫn còn trong `x`, nhưng nếu chỉ join w:t sẽ mất token.
  // Vì vậy cách chắc nhất: strip tag nhưng CHỪA token.

  // Bọc token bằng placeholder không có dấu <>
  x = x.replace(/\[!m:\$(.*?)\$\]/g, "___MATH_TOKEN_START___$1___MATH_TOKEN_END___");
  x = x.replace(/\[!img:\$(.*?)\$\]/g, "___IMG_TOKEN_START___$1___IMG_TOKEN_END___");

  // Thay w:t bằng content để giữ cả token lẫn text
  x = x.replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g, (seg) => {
    const mm = seg.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/);
    return mm ? mm[1] : "";
  });

  // Xóa toàn bộ tag còn lại
  x = x.replace(/<[^>]+>/g, "");

  // Khôi phục token
  x = x
    .replace(/___MATH_TOKEN_START___(.*?)___MATH_TOKEN_END___/g, "[!m:$$$1$$]")
    .replace(/___IMG_TOKEN_START___(.*?)___IMG_TOKEN_END___/g, "[!img:$$$1$$]");

  // Decode entity + normalize whitespace vừa phải
  x = decodeXmlEntities(x)
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return x;
}

function parseQuestions(text) {
  // bạn đang format kiểu: "Câu 1." "A." ...
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

    // NOTE: regex choices có thể gặp trường hợp nội dung chứa "A." trong text.
    // Mình giữ gần giống bạn, nhưng an toàn hơn bằng cách chặn tới (B.|C.|D.|Lời giải|$)
    const choiceRe = /(\*?)([A-D])\.\s([\s\S]*?)(?=\n\*?[A-D]\.\s|\nLời giải|$)/g;
    let m;
    while ((m = choiceRe.exec(main))) {
      const starred = m[1] === "*";
      const label = m[2];
      const content = (m[3] || "").trim();
      if (starred) q.correct = label;
      q.choices.push({ label, text: content });
    }

    // content = phần trước lựa chọn đầu tiên
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
    const relEntry = zip.files.find((f) => f.path === "word/_rels/document.xml.rels");
    if (!docEntry || !relEntry) throw new Error("Missing document.xml or document.xml.rels");

    let docXml = (await docEntry.buffer()).toString("utf8");
    const relsXml = (await relEntry.buffer()).toString("utf8");
    const rels = parseRels(relsXml);

    // A) Ảnh (drawing/vml) -> token [!img:$img_x$]
    const imgTok = await tokenizeImages(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    const imgMap = imgTok.imgMap;

    // B) MathType OLE -> token [!m:$mathtype_x$] (+ fallback ảnh nếu cần)
    const mathTok = await tokenizeMathTypeOle(docXml, rels, zip.files, imgMap);
    docXml = mathTok.outXml;
    const mathMap = mathTok.mathMap;

    // C) XML -> text giữ token
    const text = wordXmlToTextKeepTokens(docXml);

    // D) parse câu hỏi
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      math: mathMap,  // mathtype_x -> "<math ...>...</math>"
      images: imgMap, // img_x -> dataURL ; fallback_mathtype_x -> dataURL
      rawText: text,  // debug nếu cần
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
