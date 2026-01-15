import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";

const app = express();
app.use(cors());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const buffer = req.file.buffer;

    // 1️⃣ đọc document.xml dạng STREAM (rất nhanh)
    const zip = await unzipper.Open.buffer(buffer);
    const entry = zip.files.find(f => f.path === "word/document.xml");
    if (!entry) throw new Error("document.xml not found");

    let xml = (await entry.buffer()).toString("utf8");

    // 2️⃣ Token hoá MathType (KHÔNG parse XML)
    let mathIndex = 0;
    const mathMap = {};

    xml = xml.replace(/<m:oMath[\s\S]*?<\/m:oMath>/g, (m) => {
      const key = `mathtype_${++mathIndex}`;
      mathMap[key] = m;
      return `[!m:$${key}$]`;
    });

    // 3️⃣ Bóc text Word cơ bản (đủ để làm quiz)
    const text = xml
      .replace(/<[^>]+>/g, "")
      .replace(/\s+/g, " ")
      .trim();

    // 4️⃣ Tách câu hỏi (logic giống Azota)
    const rawQuestions = text.split(/(?=Câu\s+\d+\.)/);

    const questions = rawQuestions
      .filter(q => q.trim().startsWith("Câu"))
      .map((block, i) => {
        const answers = [];
        let correct = null;

        block.replace(/\*?([A-D])\.\s([^A-D]*)/g, (_, label, content) => {
          if (_.startsWith("*")) correct = label;
          answers.push({ label, text: content.trim() });
        });

        return {
          id: i + 1,
          content: block.split(/A\./)[0].trim(),
          answers,
          correct
        };
      });

    // 5️⃣ TRẢ JSON – KHÔNG RENDER – RẤT NHANH
    res.json({
      ok: true,
      questionCount: questions.length,
      questions,
      math: mathMap
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
