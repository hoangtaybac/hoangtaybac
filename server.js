import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import mime from "mime-types";

const app = express();
app.use(cors());

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

/**
 * Lưu tạm “media” trong RAM theo lần upload gần nhất (MVP).
 * Khi làm sản phẩm thật: lưu S3/R2 hoặc disk/redis.
 */
const MEM = {
  images: new Map(), // key -> { buf, mime }
};

function stripTagsKeepTokens(xml) {
  // giữ lại các token như [img:$..$] sau khi strip
  return xml.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

function parseRels(relsXml) {
  // Map rId -> target
  const map = new Map();
  const re = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  let m;
  while ((m = re.exec(relsXml))) {
    map.set(m[1], m[2]);
  }
  return map;
}

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    MEM.images.clear();

    const zip = await unzipper.Open.buffer(req.file.buffer);

    const docEntry = zip.files.find(f => f.path === "word/document.xml");
    const relEntry = zip.files.find(f => f.path === "word/_rels/document.xml.rels");
    if (!docEntry || !relEntry) throw new Error("Missing document.xml or document.xml.rels");

    let docXml = (await docEntry.buffer()).toString("utf8");
    const relsXml = (await relEntry.buffer()).toString("utf8");
    const rels = parseRels(relsXml);

    // 1) Extract ALL word/media/* into MEM (png/jpg/wmf…)
    // NOTE: WMF browser không hiển thị được. Nhưng nhiều ảnh thật là png/jpg vẫn ok.
    // Với MathType trong file bạn gửi đa phần là .wmf → bước sau sẽ convert (nâng cấp).
    for (const f of zip.files) {
      if (f.path.startsWith("word/media/")) {
        const buf = await f.buffer();
        const mt = mime.lookup(f.path) || "application/octet-stream";
        MEM.images.set(f.path.replace("word/media/", ""), { buf, mime: mt });
      }
    }

    // 2) Thay các OLE MathType (w:object có v:imagedata r:id="rId..") thành token eq_#
    // Lấy rId của v:imagedata -> rels -> target = media/imageX.wmf (hoặc png)
    let eqIndex = 0;
    const eqMap = {}; // eq_1 -> filename trong media

    docXml = docXml.replace(/<w:object[\s\S]*?<\/w:object>/g, (block) => {
      const m = block.match(/<v:imagedata\b[^>]*\br:id="([^"]+)"/);
      if (!m) return block;

      const rId = m[1];
      const target = rels.get(rId); // e.g. "media/image1.wmf"
      if (!target || !target.startsWith("media/")) return block;

      const filename = target.replace("media/", "");
      const key = `eq_${++eqIndex}`;
      eqMap[key] = filename;

      // token cho equation dạng ảnh
      return `[img:$${key}$]`;
    });

    // 3) Strip tags → text (giữ token)
    const text = stripTagsKeepTokens(docXml);

    // 4) Parse câu hỏi / đáp án / lời giải (nhanh)
    const blocks = text.split(/(?=Câu\s+\d+\.)/);
    const questions = [];
    for (const b of blocks) {
      if (!b.startsWith("Câu")) continue;
      const [main, sol] = b.split(/Lời giải/i);

      const q = { content: "", choices: [], correct: null, solution: sol ? sol.trim() : "" };

      main.replace(/\*?([A-D])\.\s([^A-D]*)/g, (mm, label, content) => {
        if (mm.startsWith("*")) q.correct = label;
        q.choices.push({ label, text: content.trim() });
        return mm;
      });

      q.content = main.split(/A\./)[0].trim();
      questions.push(q);
    }

    // 5) Trả JSON: images trả URL (client dùng luôn)
    const images = {};
    for (const [k, filename] of Object.entries(eqMap)) {
      images[k] = `${req.protocol}://${req.get("host")}/img/${encodeURIComponent(filename)}`;
    }

    res.json({
      ok: true,
      total: questions.length,
      questions,
      images,  // eq_1 -> URL ảnh
      note: "MathType trong file của bạn là OLE (oleObject*.bin) + preview .wmf. Đang hiển thị bằng ảnh."
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// Serve ảnh từ RAM theo filename trong word/media/
app.get("/img/:name", (req, res) => {
  const name = req.params.name;
  const item = MEM.images.get(name);
  if (!item) return res.status(404).send("not found");

  res.setHeader("Content-Type", item.mime);
  res.send(item.buf);
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 running on", PORT));
