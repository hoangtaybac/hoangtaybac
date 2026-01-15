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

/* ================= UTILS ================= */

function stripTags(xml) {
  return xml
    .replace(/<[^>]+>/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function extractMath(xml) {
  let index = 0;
  const map = {};
  xml = xml.replace(/<m:oMath[\s\S]*?<\/m:oMath>/g, m => {
    const key = `mathtype_${++index}`;
    map[key] = m;
    return `[!m:$${key}$]`;
  });
  return { xml, map };
}

function extractImages(xml) {
  let index = 0;
  const map = {};
  xml = xml.replace(/<w:drawing[\s\S]*?<\/w:drawing>/g, m => {
    const key = `img_${++index}`;
    map[key] = m;
    return `[img:$${key}$]`;
  });
  return { xml, map };
}

/* =============== PARSE QUIZ =============== */

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
      solution: ""
    };

    // Lời giải
    const [main, solution] = block.split(/Lời giải/i);
    q.solution = solution ? solution.trim() : "";

    // Đáp án
    main.replace(/\*?([A-D])\.\s([^A-D]*)/g, (m, label, text) => {
      if (m.startsWith("*")) q.correct = label;
      q.choices.push({ label, text: text.trim() });
    });

    // Nội dung câu hỏi
    q.content = main.split(/A\./)[0].trim();

    questions.push(q);
  }

  return questions;
}

/* ================= API ================= */

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const zip = await unzipper.Open.buffer(req.file.buffer);
    const doc = zip.files.find(f => f.path === "word/document.xml");
    if (!doc) throw new Error("document.xml not found");

    let xml = (await doc.buffer()).toString("utf8");

    // 1️⃣ MathType → token
    const mathRes = extractMath(xml);
    xml = mathRes.xml;

    // 2️⃣ Image → token
    const imgRes = extractImages(xml);
    xml = imgRes.xml;

    // 3️⃣ Text
    const text = stripTags(xml);

    // 4️⃣ Parse quiz
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      math: mathRes.map,
      images: imgRes.map
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("🚀 Quiz Engine running on port", PORT);
});
