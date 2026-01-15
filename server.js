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

function stripTags(xml) {
  return xml
    .replace(/<[^>]+>/g, "")
    .replace(/\s+/g, " ")
    .trim();
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

/**
 * Try extract embedded MathML from MathType OLE binary.
 * MathType often stores MathML translation inside the object :contentReference[oaicite:1]{index=1}
 */
function extractMathMLFromOle(buf) {
  // Try UTF-8 scan
  const utf8 = buf.toString("utf8");
  let i = utf8.indexOf("<math");
  if (i !== -1) {
    let j = utf8.indexOf("</math>", i);
    if (j !== -1) return utf8.slice(i, j + 7);
  }

  // Try UTF-16LE scan (common in OLE)
  const u16 = buf.toString("utf16le");
  i = u16.indexOf("<math");
  if (i !== -1) {
    let j = u16.indexOf("</math>", i);
    if (j !== -1) return u16.slice(i, j + 7);
  }

  // Some store xmlns on <math ...>, still ok
  return null;
}

/**
 * Replace <w:object ...> blocks with [!m:$mathtype_x$]
 * Use v:imagedata r:id and o:OLEObject r:id to locate oleObject bin via rels.
 */
function tokenizeMathTypeOle(docXml, rels, zipFiles) {
  let idx = 0;
  const mathMap = {}; // mathtype_1 -> MathML string

  const getZipEntryBuffer = async (path) => {
    const f = zipFiles.find((x) => x.path === path);
    if (!f) return null;
    return await f.buffer();
  };

  // Replace each w:object block
  // IMPORTANT: In Word, MathType OLE often looks like:
  // <w:object> ... <o:OLEObject r:id="rIdX" .../> ... <v:imagedata r:id="rIdY" .../>
  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;

  const jobs = [];
  let out = docXml.replace(OBJECT_RE, (block) => {
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid); // e.g. "embeddings/oleObject1.bin"
    if (!oleTarget) return block;

    const key = `mathtype_${++idx}`;

    // create async job to extract MathML
    jobs.push(
      (async () => {
        const fullPath = oleTarget.startsWith("embeddings/")
          ? `word/${oleTarget}`
          : `word/${oleTarget}`;

        const buf = await getZipEntryBuffer(fullPath);
        if (!buf) return;

        const mml = extractMathMLFromOle(buf);
        if (mml) {
          mathMap[key] = mml;
        } else {
          // fallback: if no MathML found, keep empty -> client will show nothing
          // (Nếu bạn muốn fallback ảnh, mình sẽ thêm ở bước sau)
          mathMap[key] = "";
        }
      })()
    );

    return `[!m:$${key}$]`;
  });

  return { outXml: out, mathMap, jobs };
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

    // 1) Tokenize MathType OLE -> [!m:$mathtype_x$]
    const { outXml, mathMap, jobs } = tokenizeMathTypeOle(docXml, rels, zip.files);
    docXml = outXml;

    // Wait extract MathML jobs
    await Promise.all(jobs);

    // 2) Strip tags -> text (keep tokens)
    const text = stripTags(docXml);

    // 3) Parse quiz blocks
    const questions = parseQuestions(text);

    res.json({
      ok: true,
      total: questions.length,
      questions,
      math: mathMap, // mathtype_x -> "<math ...>...</math>"
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
