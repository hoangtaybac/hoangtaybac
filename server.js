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
 * - mở rộng để phát hiện MathML bị escape (&lt;math ... &gt;)
 */
function extractMathMLFromOleScan(buf) {
  try {
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

    // Thử tìm MathML bị escape như &lt;math ... &gt;
    i = utf8.indexOf("&lt;math");
    if (i !== -1) {
      let j = utf8.indexOf("&lt;/math&gt;", i);
      if (j !== -1) {
        const esc = utf8.slice(i, j + "&lt;/math&gt;".length);
        return decodeXmlEntities(esc);
      }
    }

    i = u16.indexOf("&lt;math");
    if (i !== -1) {
      let j = u16.indexOf("&lt;/math&gt;", i);
      if (j !== -1) {
        const esc = u16.slice(i, j + "&lt;/math&gt;".length);
        return decodeXmlEntities(esc);
      }
    }

    return null;
  } catch {
    return null;
  }
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

    // 1) thử convert bằng thư viện chính
    try {
      const latex = (MathMLToLaTeX.convert(mml) || "").trim();
      if (latex) return latex;
    } catch (e) {
      // nếu thư viện ném lỗi thì tiếp tục fallback
      console.debug("MathMLToLaTeX.convert error:", e?.message || e);
    }

    // 2) nếu không có kết quả, thử xử lý một vài tag MathML thông dụng thủ công
    let s = mml;

    // a) xử lý <msqrt>...</msqrt> -> \sqrt{...}
    s = s.replace(/<msqrt\b[^>]*>([\s\S]*?)<\/msqrt>/gi, (m, inner) => {
      const content = inner.replace(/<[^>]+>/g, "").trim();
      // trả về LaTeX dạng \sqrt{...}
      return `\\sqrt{${content}}`;
    });

    // b) xử lý <mroot>radicand index</mroot> -> \sqrt[index]{radicand}
    s = s.replace(/<mroot\b[^>]*>([\s\S]*?)<\/mroot>/gi, (m, inner) => {
      // Thử tách bằng heuristic: phần đầu là radicand, phần sau là index
      // Bằng cách tìm phần tử con cuối cùng là index
      // Loại bỏ tag, split theo thẻ đóng -> mở
      const parts = inner
        .replace(/>\s+</g, "><")
        .split(/<\/[^>]+>\s*<[^>]+>/)
        .map((p) => p.replace(/<[^>]+>/g, "").trim())
        .filter(Boolean);
      if (parts.length >= 2) {
        const rad = parts[0];
        const idx = parts[1];
        return `\\sqrt[${idx}]{${rad}}`;
      }
      // fallback: strip tags
      const stripped = inner.replace(/<[^>]+>/g, "").trim();
      return stripped;
    });

    // c) loại bỏ tag còn lại và decode entity để lấy text plain (heuristic)
    s = decodeXmlEntities(s.replace(/<[^>]+>/g, "")).trim();

    // d) nếu sau xử lý còn text thì coi đó là LaTeX tạm thời
    if (s) return s;

    return "";
  } catch (e) {
    console.debug("mathmlToLatexSafe error:", e?.message || e);
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
      if (oleBuf) {
        try {
          mml = extractMathMLFromOleScan(oleBuf) || "";
        } catch (e) {
          mml = "";
        }
      }

      if (mml) {
        console.debug(`[MathType] found MathML directly for ${key} (len=${mml.length})`);
      } else {
        console.debug(`[MathType] no direct MathML found for ${key}`);
      }

      // 2) fallback ruby convert (MTEF inside OLE)
      if (!mml && oleBuf) {
        try {
          const out = await rubyOleToMathML(oleBuf);
          if (out) {
            mml = out;
            console.debug(`[MathType] ruby mt2mml produced MathML for ${key} (len=${mml.length})`);
          } else {
            console.debug(`[MathType] ruby mt2mml produced empty output for ${key}`);
          }
        } catch (e) {
          console.debug(`[MathType] ruby mt2mml failed for ${key}:`, e?.message || e);
          mml = "";
        }
      }

      // 3) MathML -> LaTeX
      const latex = mml ? mathmlToLatexSafe(mml) : "";
      if (latex) {
        latexMap[key] = latex;
        console.debug(`[MathType] latex for ${key}:`, latex.slice(0, 200));
        return;
      }

      if (mml && !latex) {
        // debug snippet
        console.debug(
          `[MathType] mathml->latex returned empty for ${key} — MathML snippet:`,
          (mml || "").slice(0, 400).replace(/\n/g, " ")
        );
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
                  console.debug(`[MathType] used converted PNG fallback for ${key}`);
                  return;
                }
              } catch (e) {
                console.debug(`[MathType] emf/wmf->png conversion failed for ${key}:`, e?.message || e);
              }
            }
            images[`fallback_${key}`] =
              `data:${mime};base64,${imgBuf.toString("base64")}`;
            console.debug(`[MathType] used image fallback for ${key} (mime=${mime})`);
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
          } catch (e) {
            console.debug("EMF/WMF conversion error:", e?.message || e);
          }
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
 * ✅ HOÀN THIỆN:
 * - GIỮ token [!m:$...$]/[!img:$...$] để HTML render công thức + ảnh
 * - GIỮ underline bằng <u>...</u>
 * - Không đổi thuật toán MathType/LaTeX phía trên
 */
function wordXmlToTextKeepTokens(docXml) {
  let x = docXml
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  // 1) Protect tokens BEFORE stripping tags (both $key$ and $$key$$)
  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

  // 2) Convert each run <w:r> while preserving underline + token text
  x = x.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
    const hasU =
      /<w:u\b[^>]*\/>/.test(run) &&
      !/<w:u\b[^>]*w:val="none"[^>]*\/>/.test(run);

    // remove run properties but keep content (tokens are plain text)
    let inner = run.replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g, "");

    // w:t -> raw text
    inner = inner.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (_, t) => t ?? "");

    // optional instrText
    inner = inner.replace(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g,
      (_, t) => t ?? ""
    );

    // remove remaining tags inside run
    inner = inner.replace(/<[^>]+>/g, "");

    if (!inner) return "";
    return hasU ? `<u>${inner}</u>` : inner;
  });

  // 3) Remove remaining tags outside runs, but keep <u>
  x = x.replace(/<(?!\/?u\b)[^>]+>/g, "");

  // 4) Restore tokens in a stable form (use $$ ... $$)
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

    // ✅ giữ nguyên thuật toán parse đáp án của bạn
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

    // 1) MathType -> LaTeX (and fallback images) ✅ giữ nguyên
    const images = {};
    const mt = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mt.outXml;
    const latexMap = mt.latexMap;

    // 2) normal images ✅ giữ nguyên
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // 3) text ✅ giờ giữ token + underline
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) parse questions ✅ giữ nguyên
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
