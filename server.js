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
  return String(s)
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

// ✅ debug/nhận diện căn thức trong MathML
const SQRT_MATHML_RE = /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

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
 * ✅ HOÀN THIỆN: cắt đúng block <math>...</math> (nếu ruby in kèm log)
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

        const out = String(stdout || "");
        const decoded = decodeXmlEntities(out);

        // cố gắng lấy đúng block MathML nếu có (kể cả namespace mml:math)
        const m =
          decoded.match(/<[^<]{0,20}math[\s\S]*?<\/[^<]{0,20}math>/i) ||
          out.match(/<[^<]{0,20}math[\s\S]*?<\/[^<]{0,20}math>/i);

        resolve(String(m ? m[0] : out).trim());
      }
    );
  });
}

/**
 * ✅ FIX CĂN THỨC: decode entity trước khi MathMLToLaTeX.convert()
 */
function mathmlToLatexSafe(mml) {
  try {
    if (!mml) return "";

    const normalized = decodeXmlEntities(String(mml).trim());

    // một số output có namespace <mml:math...>
    if (!normalized.includes("<math") && !normalized.includes(":math"))
      return "";

    const latex = MathMLToLaTeX.convert(normalized);
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
 * ✅ HOÀN THIỆN: trả thêm sqrtDebug để debug căn thức
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
  const sqrtDebug = []; // ✅ debug căn thức

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

      // ✅ debug căn thức: check trước khi convert
      const mmlDecoded = decodeXmlEntities(mml || "");
      const hasSqrt =
        SQRT_MATHML_RE.test(mml || "") || SQRT_MATHML_RE.test(mmlDecoded);

      // 3) MathML -> LaTeX (đã decode trong mathmlToLatexSafe)
      const latex = mml ? mathmlToLatexSafe(mml) : "";

      if (hasSqrt) {
        sqrtDebug.push({
          key,
          ole: oleFull,
          hasSqrt,
          mmlSnippet: (mmlDecoded || "").slice(0, 1200),
          latex,
          ok: Boolean(latex),
        });
      }

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

  return { outXml: docXml, latexMap, sqrtDebug };
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

    // 1) MathType -> LaTeX (and fallback images) ✅ giữ nguyên flow, thêm debug căn thức
    const images = {};
    const mt = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mt.outXml;
    const latexMap = mt.latexMap;
    const sqrtDebug = mt.sqrtDebug || [];

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
      sqrtDebug, // ✅ thêm debug căn thức
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
// ĐOẠN NÀY 
/* ================== SPACE AROUND MATH ================== */
function lastVisibleChar(html) {
  const t = String(html || "").replace(/<[^>]*>/g, "");
  return t.length ? t[t.length - 1] : "";
}

function appendMathWithOneSpace(html, mathSpan) {
  const prev = lastVisibleChar(html);
  if (prev && !/\s/.test(prev)) html += " ";
  html += mathSpan;
  html += " ";
  return html;
}

/* ================== RENDER PARAGRAPH / TABLE ================== */
function renderParagraph(pNode, ctx) {
  const { latexByRid, imageByRid, debug } = ctx;
  let html = "";

  const runs = kids(pNode, "w:r");
  for (const rWrap of runs) {
    const rNode = rWrap["w:r"];
    const under = runIsUnderlined(rNode);

    if (Array.isArray(rNode)) {
      for (const child of rNode) {
        if (child["w:t"]) {
          const text = getTextFromPreserveWrap(child, "w:t");
          if (text) {
            const esc = escapeTextToHtml(text);
            html += under ? `<u>${esc}</u>` : esc;
          }
        }

        if (child["w:tab"]) html += "&emsp;";
        if (child["w:br"]) html += "<br/>";

        if (child["a:blip"] || child["pic:blipFill"] || child["w:drawing"]) {
          const imgRids = unique(findImageEmbedRidsDeep(child, []));
          for (const rid of imgRids) {
            const dataUri = imageByRid[rid];
            if (dataUri) {
              debug.imagesInjected++;
              html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
            }
          }
        }

        if (child["w:pict"] || child["v:shape"]) {
          const imgRids = unique(findImageEmbedRidsDeep(child, []));
          for (const rid of imgRids) {
            const dataUri = imageByRid[rid];
            if (dataUri) {
              debug.imagesInjected++;
              html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
            }
          }
        }

        if (child["w:object"] || child["o:OLEObject"]) {
          const allRids = unique(findAllRidsDeep(child, []));

          let foundMath = false;
          for (const rid of allRids) {
            const latex = latexByRid[rid];
            if (latex) {
              debug.seenOle++;
              if (debug.sampleRids.length < 12) debug.sampleRids.push(rid);
              const mathSpan = `<span class="math">\\(${latex}\\)</span>`;
              html = appendMathWithOneSpace(html, mathSpan);
              debug.oleInjected++;
              foundMath = true;
            }
          }

          // fallback to image if no latex
          if (!foundMath) {
            const imgRids = unique(findImageEmbedRidsDeep(child, []));
            for (const rid of imgRids) {
              const dataUri = imageByRid[rid];
              if (dataUri) {
                debug.imagesInjected++;
                html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
              }
            }
          }
        }
      }
    } else {
      const runText = collectTextFromRun(rNode);
      if (runText) {
        const esc = escapeTextToHtml(runText);
        html += under ? `<u>${esc}</u>` : esc;
      }
    }

    const runImgRids = unique(findImageEmbedRidsDeep(rNode, []));
    const processedInLoop = new Set();

    if (Array.isArray(rNode)) {
      for (const child of rNode) {
        if (child["w:drawing"] || child["w:pict"] || child["v:shape"] || child["w:object"]) {
          const childRids = findImageEmbedRidsDeep(child, []);
          childRids.forEach((rid) => processedInLoop.add(rid));
        }
      }
    }

    for (const rid of runImgRids) {
      if (processedInLoop.has(rid)) continue;
      const dataUri = imageByRid[rid];
      if (dataUri) {
        debug.imagesInjected++;
        html += `<img src="${dataUri}" style="max-width:100%;height:auto;vertical-align:middle;" />`;
      }
    }

    if (runHasOleLike(rNode)) {
      debug.seenOleRuns++;
      const rids = unique(findAllRidsDeep(rNode, []));

      const processedMathRids = new Set();
      if (Array.isArray(rNode)) {
        for (const child of rNode) {
          if (child["w:object"] || child["o:OLEObject"]) {
            const childRids = findAllRidsDeep(child, []);
            childRids.forEach((rid) => {
              if (latexByRid[rid]) processedMathRids.add(rid);
            });
          }
        }
      }

      for (const rid of rids) {
        if (processedMathRids.has(rid)) continue;
        const latex = latexByRid[rid];
        if (latex) {
          debug.seenOle++;
          if (debug.sampleRids.length < 12) debug.sampleRids.push(rid);
          const mathSpan = `<span class="math">\\(${latex}\\)</span>`;
          html = appendMathWithOneSpace(html, mathSpan);
          debug.oleInjected++;
        } else {
          debug.ignoredRids++;
        }
      }
    }
  }

  return html;
}

function renderTable(tblNode, ctx) {
  const rows = kids(tblNode, "w:tr");
  let html = `<table border="1" style="border-collapse:collapse;width:auto;max-width:100%;">`;

  for (const trWrap of rows) {
    const trNode = trWrap["w:tr"];
    html += "<tr>";

    const cells = kids(trNode, "w:tc");
    for (const tcWrap of cells) {
      const tcNode = tcWrap["w:tc"];
      html += `<td style="padding:6px;vertical-align:top;">`;

      const paras = kids(tcNode, "w:p");
      for (const pWrap of paras) {
        const pHtml = renderParagraph(pWrap["w:p"], ctx);
        if (pHtml) html += pHtml;
        html += "<br/>";
      }

      html += "</td>";
    }

    html += "</tr>";
  }

  html += "</table><br/>";
  return html;
}

function buildInlineHtml(documentXml, ctx) {
  const parser = new XMLParser({ ignoreAttributes: false, preserveOrder: true });
  const tree = parser.parse(documentXml);
  const doc = kids(tree, "w:document")[0]?.["w:document"];
  const body = kids(doc, "w:body")[0]?.["w:body"];
  const bodyChildren = Array.isArray(body) ? body : [];

  let html = "";
  for (const child of bodyChildren) {
    if (child["w:p"]) {
      const pHtml = renderParagraph(child["w:p"], ctx);
      if (pHtml) html += pHtml;
      html += "<br/>";
    } else if (child["w:tbl"]) {
      html += renderTable(child["w:tbl"], ctx);
    }
  }
  return html;
}

/* ================== FORMAT LAYOUT ================== */
function splitByMath(html) {
  const out = [];
  const re = /\\\([\s\S]*?\\\)/g;
  let last = 0,
    m;
  while ((m = re.exec(html)) !== null) {
    if (m.index > last) out.push({ math: false, text: html.slice(last, m.index) });
    out.push({ math: true, text: m[0] });
    last = m.index + m[0].length;
  }
  if (last < html.length) out.push({ math: false, text: html.slice(last) });
  return out;
}

function normalizeGluedChoiceMarkers(s) {
  s = String(s || "");
  s = s.replace(/([^<\s>])([ABCD])\./g, "$1 $2.");
  s = s.replace(/([^<\s>])([a-d])\)/gi, "$1 $2)");
  s = s.replace(/([^<\s>])(<u[^>]*>\s*[ABCD]\s*<\/u>\s*\.)/gi, "$1 $2");
  s = s.replace(/([^<\s>])(<u[^>]*>\s*[a-d]\s*<\/u>\s*\))/gi, "$1 $2");
  return s;
}

function formatAbcdOutsideHeaders(text) {
  const headerRegex = /(<div class="section-header">[\s\S]*?<\/div>)/g;
  const segments = text.split(headerRegex);

  return segments
    .map((seg) => {
      if (seg.startsWith('<div class="section-header">')) return seg;

      let s = seg;

      s = s
        .replace(/(^|<br\/>\s*<br\/>|\n)\s*([a-d])\)/gi, "$1&emsp;$2)")
        .replace(/([^<\n])\s*([a-d])\)/gi, "$1<br/>&emsp;$2)");

      s = s
        .replace(/(^|<br\/>\s*<br\/>|\n)\s*(<u[^>]*>\s*[a-d]\s*\)\s*<\/u>)/gi, "$1&emsp;$2")
        .replace(/([^<\n])\s*(<u[^>]*>\s*[a-d]\s*\)\s*<\/u>)/gi, "$1<br/>&emsp;$2");

      s = s
        .replace(/(^|<br\/>\s*<br\/>|\n)\s*(<u[^>]*>\s*[a-d]\s*<\/u>\s*\))/gi, "$1&emsp;$2")
        .replace(/([^<\n])\s*(<u[^>]*>\s*[a-d]\s*<\/u>\s*\))/gi, "$1<br/>&emsp;$2");

      return s;
    })
    .join("");
}

function formatExamLayout(html) {
  let result = html;

  result = result.replace(/\s+/g, " ");
  result = result.replace(/PHẦN(\d)/gi, "PHẦN $1");

  result = result.replace(
    /(^|<br\/>)\s*(PHẦN\s+\d+\.(?:(?!<br\/>\s*Câu\s+\d).)*)/g,
    "$1<br/><div class=\"section-header\"><strong>$2</strong></div>"
  );

  const parts = splitByMath(result);

  for (const p of parts) {
    if (p.math) continue;

    p.text = normalizeGluedChoiceMarkers(p.text);

    p.text = p.text
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*([ABCD])\./g, "$1&emsp;$2.")
      .replace(/([^<\n])\s*([ABCD])\./g, "$1<br/>&emsp;$2.");

    p.text = p.text
      .replace(/(^|<br\/>\s*<br\/>|\n)\s*(<u[^>]*>\s*[ABCD]\s*<\/u>\s*\.)/gi, "$1&emsp;$2")
      .replace(/([^<\n])\s*(<u[^>]*>\s*[ABCD]\s*<\/u>\s*\.)/gi, "$1<br/>&emsp;$2");

    p.text = formatAbcdOutsideHeaders(p.text);

    p.text = p.text.replace(/(Câu)\s*(\d+)\s*\./g, "$1 $2.");
    p.text = p.text.replace(/(<br\/>\s*){3,}/g, "<br/><br/>");
  }

  return parts.map((x) => x.text).join("");
}

/* ================== EXAM PARSING + SOLUTION SPLIT ================== */
function stripAllTagsToPlain(html) {
  return String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/&emsp;/g, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function detectHasMCQ(plain) {
  const marks = plain.match(/\b[ABCD]\./g) || [];
  return new Set(marks).size >= 2;
}

function detectHasTF4(plain) {
  const marks = plain.match(/\b[a-d]\)/gi) || [];
  return new Set(marks.map((x) => x.toLowerCase())).size >= 2;
}

function findSolutionMarkerIndex(html, fromIndex = 0) {
  const s = String(html || "");
  const re =
    /(Lời(?:\s*<[^>]*>)*\s*giải|Giải(?:\s*<[^>]*>)*\s*chi\s*ti\s*ết|Hướng(?:\s*<[^>]*>)*\s*dẫn(?:\s*<[^>]*>)*\s*giải)/i;
  const sub = s.slice(fromIndex);
  const m = re.exec(sub);
  if (!m) return -1;
  return fromIndex + m.index;
}

function splitSolutionSections(tailHtml) {
  let s = String(tailHtml || "").trim();
  if (!s) return { solutionHtml: "", detailHtml: "" };

  const reCT = /(Giải(?:\s*<[^>]*>)*\s*chi\s*ti\s*ết)/i;
  const matchCT = reCT.exec(s);

  if (matchCT) {
    const idxCT = matchCT.index;
    return {
      solutionHtml: s.slice(0, idxCT).trim(),
      detailHtml: s.slice(idxCT).trim(),
    };
  }

  return { solutionHtml: s, detailHtml: "" };
}

function extractUnderlinedKeys(blockHtml) {
  const keys = { mcq: null, tf: [] };
  const s = String(blockHtml || "");

  let m =
    s.match(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./i) ||
    s.match(/<u[^>]*>\s*([A-D])\.\s*<\/u>/i);
  if (m) keys.mcq = m[1].toUpperCase();

  let mm;
  const reTF1 = /<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi;
  while ((mm = reTF1.exec(s)) !== null) keys.tf.push(mm[1].toLowerCase());

  const reTF2 = /<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi;
  while ((mm = reTF2.exec(s)) !== null) keys.tf.push(mm[1].toLowerCase());

  keys.tf = [...new Set(keys.tf)];
  return keys;
}

function normalizeUnderlinedMarkersForSplit(html) {
  let s = String(html || "");
  s = s.replace(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./gi, "$1.");
  s = s.replace(/<u[^>]*>\s*([A-D])\.\s*<\/u>/gi, "$1.");
  s = s.replace(/<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi, "$1)");
  s = s.replace(/<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi, "$1)");
  return s;
}

function removeUnsupportedImages(html) {
  let s = String(html || "");

  s = s.replace(/<img[^>]*src\s*=\s*["']\s*["'][^>]*>/gi, "");
  s = s.replace(/<img(?![^>]*src\s*=)[^>]*>/gi, "");
  s = s.replace(/<img[^>]*data:application\/octet-stream[^>]*>/gi, "");

  return s;
}

function splitChoicesHtmlABCD(blockHtml) {
  let s = normalizeUnderlinedMarkersForSplit(blockHtml);
  s = s.replace(/&emsp;/g, " ");
  s = normalizeGluedChoiceMarkers(s);
  s = s.replace(/<br\/>/g, " <br/>");

  const re = /(^|[\s>.:;,<\)\]\}！？\?])([ABCD])\./g;

  const hits = [];
  let m;
  while ((m = re.exec(s)) !== null) hits.push({ idx: m.index + m[1].length, key: m[2] });
  if (hits.length < 2) return null;

  const lastStart = hits[hits.length - 1].idx;
  const solIdx = findSolutionMarkerIndex(s, lastStart);
  const endAll = solIdx >= 0 ? solIdx : s.length;

  const out = {
    _stem: s.slice(0, hits[0].idx).trim(),
    _tail: solIdx >= 0 ? s.slice(solIdx).trim() : "",
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : endAll;
    let seg = s.slice(start, end).trim();
    seg = seg.replace(/^([ABCD])\.\s*/i, "");

    out[key] = removeUnsupportedImages(seg.trim());
  }
  return out;
}

function splitStatementsHtmlabcd(blockHtml) {
  let s = normalizeUnderlinedMarkersForSplit(blockHtml);
  s = s.replace(/&emsp;/g, " ");
  s = normalizeGluedChoiceMarkers(s);
  s = s.replace(/<br\/>/g, " <br/>");

  const earlysolIdx = findSolutionMarkerIndex(s, 0);
  let workingHtml = s;
  let tailHtml = "";
  
  if (earlysolIdx >= 0) {
    workingHtml = s.slice(0, earlysolIdx);
    tailHtml = s.slice(earlysolIdx).trim();
  }

  const re = /(^|[\s>.:;,<\)\]\}！？\?])([a-d])\)/gi;

  const hits = [];
  let m;
  while ((m = re.exec(workingHtml)) !== null) {
    hits.push({ idx: m.index + m[1].length, key: m[2].toLowerCase() });
  }
  if (hits.length < 2) return null;

  const out = {
    _stem: workingHtml.slice(0, hits[0].idx).trim(),
    _tail: tailHtml,
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : workingHtml.length;
    let seg = workingHtml.slice(start, end).trim();
    seg = seg.replace(/^([a-d])\)\s*/i, "");
    
    out[key] = removeUnsupportedImages(seg.trim());
  }
  return out;
}

function cleanStem(html) {
  if (!html) return html;
  return String(html).replace(/^Câu\s+\d+\.?\s*/i, '').trim();
}

function parseExamFromInlineHtml(inlineHtml) {
  const re = /(^|<br\/>\s*)\s*(?:<[^>]*>\s*)*Câu\s+(\d+)\./gi;

  const hits = [];
  let m;
  while ((m = re.exec(inlineHtml)) !== null) {
    const startAt = m.index + m[1].length;
    hits.push({ qno: Number(m[2]), pos: startAt });
  }
  if (!hits.length) return null;

  const sectionRe = /<div class="section-header"><strong>([\s\S]*?)<\/strong><\/div>/gi;
  const sections = [];
  let sectionMatch;
  while ((sectionMatch = sectionRe.exec(inlineHtml)) !== null) {
    sections.push({
      pos: sectionMatch.index,
      html: sectionMatch[0],
      title: sectionMatch[1].trim()
    });
  }

  const rawBlocks = [];
  for (let i = 0; i < hits.length; i++) {
    const start = hits[i].pos;
    let end = i + 1 < hits.length ? hits[i + 1].pos : inlineHtml.length;
    
    for (const sec of sections) {
      if (sec.pos > start && sec.pos < end) {
        end = sec.pos;
        break;
      }
    }
    
    rawBlocks.push({ qno: hits[i].qno, pos: hits[i].pos, html: inlineHtml.slice(start, end) });
  }

  const blocks = [];
  for (const b of rawBlocks) {
    const last = blocks[blocks.length - 1];
    if (last && last.qno === b.qno) {
      last.html += "<br/>" + b.html;
    } else {
      blocks.push({ ...b });
    }
  }

  const exam = { version: 8, questions: [], sections };

  function findSectionForQuestion(qPos) {
    let currentSection = null;
    for (const sec of sections) {
      if (sec.pos < qPos) {
        currentSection = sec;
      } else {
        break;
      }
    }
    return currentSection;
  }

  for (const b of blocks) {
    const under = extractUnderlinedKeys(b.html);
    const plain = stripAllTagsToPlain(b.html);
    
    const section = findSectionForQuestion(b.pos);

    const isMCQ = detectHasMCQ(plain);
    const isTF4 = !isMCQ && detectHasTF4(plain);

    if (isMCQ) {
      const parts = splitChoicesHtmlABCD(b.html);
      const sol = splitSolutionSections(parts?._tail || "");
      exam.questions.push({
        no: b.qno,
        type: "mcq",
        stemHtml: cleanStem(parts?._stem || b.html),
        choicesHtml: { A: parts?.A || "", B: parts?.B || "", C: parts?.C || "", D: parts?.D || "" },
        answer: under.mcq,
        solutionHtml: sol.solutionHtml,
        detailHtml: sol.detailHtml,
        _plain: plain,
        section: section ? { title: section.title, html: section.html } : null
      });
      continue;
    }

    if (isTF4) {
      const parts = splitStatementsHtmlabcd(b.html);
      const sol = splitSolutionSections(parts?._tail || "");

      const ans = { a: null, b: null, c: null, d: null };
      for (const k of ["a", "b", "c", "d"]) {
        if (under.tf.includes(k)) ans[k] = true;
      }

      exam.questions.push({
        no: b.qno,
        type: "tf4",
        stemHtml: cleanStem(parts?._stem || b.html),
        statements: { a: parts?.a || "", b: parts?.b || "", c: parts?.c || "", d: parts?.d || "" },
        answer: ans,
        solutionHtml: sol.solutionHtml,
        detailHtml: sol.detailHtml,
        _plain: plain,
        section: section ? { title: section.title, html: section.html } : null
      });
      continue;
    }

    const solIdx = findSolutionMarkerIndex(b.html, 0);
    const stemPart = solIdx >= 0 ? b.html.slice(0, solIdx).trim() : b.html;
    const tailPart = solIdx >= 0 ? b.html.slice(solIdx).trim() : "";

    const sol = splitSolutionSections(tailPart);

    exam.questions.push({
      no: b.qno,
      type: "short",
      stemHtml: cleanStem(stemPart),
      boxes: 4,
      solutionHtml: sol.solutionHtml || tailPart, 
      detailHtml: sol.detailHtml || "",
      _plain: plain,
      section: section ? { title: section.title, html: section.html } : null
    });
  }

  return exam;
}

/* ================== ROUTES ================== */
app.get("/", (req, res) => {
  res.type("text").send("MathType Converter API: POST /convert-docx-html, GET /health");
});

app.get("/health", (req, res) => {
  res.json({ ok: true, node: process.version, cwd: process.cwd() });
});

app.post("/convert-docx-html", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) {
      return res.status(400).json({ ok: false, error: "No file uploaded. Field name must be 'file'." });
    }

    res.setHeader('Content-Type', 'application/json; charset=utf-8');

    const zip = await openDocxZip(req.file.buffer);

    const docBuf = await readZipEntry(zip, "word/document.xml");
    const relBuf = await readZipEntry(zip, "word/_rels/document.xml.rels");
    if (!docBuf || !relBuf) {
      return res.status(400).json({ ok: false, error: "Missing word/document.xml or word/_rels/document.xml.rels" });
    }

    const { emb: embRelMap, media: mediaRelMap } = buildRelMaps(relBuf.toString("utf8"));

    const latexByRid = {};
    const mathmlByRid = {};  // ✅ NEW: Store original MathML for debugging
    let latexOk = 0;
    let latexSanitized = 0;
    
    // ✅ Debug info for sqrt (and "interesting" MathML sampling)
    const SQRT_RE = SQRT_MATHML_RE;

    const sqrtDebug = {
      mathmlWithSqrt: 0,
      latexWithSqrt: 0,
      samples: [],
      moTokens: {}
    };

    function pickInterestingMathml(mathml) {
      const s = String(mathml || "");
      if (SQRT_RE.test(s)) return true;

      // Collect <mo> tokens to find non-trivial operators (often includes √-like glyphs)
      const moTokens = [...s.matchAll(/<mo>\s*([^<]{1,24})\s*<\/mo>/gi)].map(m => m[1].trim());
      return moTokens.some(t => {
        if (!t) return false;
        // ignore common punctuation/operators
        if (/^[\.,\+\-\=\(\)\[\]\{\}\|:;<>\/\\]$/.test(t)) return false;
        // non-ascii or entity-like
        return /[^\x00-\x7F]/.test(t) || t.includes("&") || t.includes("#");
      });
    }

    for (const [rid, embPath] of Object.entries(embRelMap)) {
      const emb = (zip.files || []).find((f) => f.path === embPath);
      if (!emb) continue;

      const buf = await emb.buffer();
      const mathml = rubyConvertOleBinToMathML(buf, embPath);
      if (!mathml) continue;

      // ✅ Store original MathML for debugging
      mathmlByRid[rid] = mathml;

      // ✅ DEBUG: collect <mo> tokens and pick interesting MathML samples
      for (const mm of mathml.matchAll(/<mo>\s*([^<]{1,24})\s*<\/mo>/gi)) {
        const tok = (mm[1] || "").trim();
        if (!tok) continue;
        sqrtDebug.moTokens[tok] = (sqrtDebug.moTokens[tok] || 0) + 1;
      }

      const hasSqrtInMathML = SQRT_RE.test(mathml);
      if (hasSqrtInMathML) sqrtDebug.mathmlWithSqrt++;

      // Save a few "interesting" MathML samples for inspection (not just the first ones)
      if (sqrtDebug.samples.length < 12 && pickInterestingMathml(mathml)) {
        sqrtDebug.samples.push({ rid, mathmlHead: mathml.slice(0, 1200) });
      }

      let latex = mathmlToLatexSafe(mathml);
      if (!latex) continue;

      const before = latex;

      latex = sanitizeLatexStrict(latex);
      latex = normalizeLatexCommands(latex);
      latex = restoreArrowAndCoreCommands(latex);
      latex = fixPiecewiseFunction(latex);
      
      // ✅ Apply sqrt post-processing again after all other fixes
      latex = postprocessLatexSqrt(latex);
      
      // ✅ Apply final cleanup (Unicode, malformed fences, spaced functions)
      latex = finalLatexCleanup(latex);

      if (latex !== before) latexSanitized++;
      
      // ✅ DEBUG: Check for sqrt in final LaTeX
      if (latex.includes('\\sqrt')) {
        sqrtDebug.latexWithSqrt++;
      }

      latexByRid[rid] = latex;
      latexOk++;
    }

    const imageByRid = {};
    let imagesOk = 0;
    let imagesConverted = 0;

    for (const [rid, mediaPath] of Object.entries(mediaRelMap)) {
      const mf = (zip.files || []).find((f) => f.path === mediaPath);
      if (!mf) continue;
      
      const buf = await mf.buffer();
      const ext = getExtFromPath(mediaPath);
      
      if (ext === "emf" || ext === "wmf") {
        const pngBuf = convertEmfWmfToPng(buf, ext);
        if (pngBuf) {
          const b64 = pngBuf.toString("base64");
          imageByRid[rid] = `data:image/png;base64,${b64}`;
          imagesConverted++;
          imagesOk++;
        } else {
          console.warn(`[SKIP_IMAGE] Failed to convert ${ext}: ${mediaPath}`);
        }
      } else {
        const mime = mimeFromExt(mediaPath);
        const b64 = buf.toString("base64");
        imageByRid[rid] = `data:${mime};base64,${b64}`;
        imagesOk++;
      }
    }

    const debug = {
      embeddings: Object.keys(embRelMap).length,
      latexCount: Object.keys(latexByRid).length,
      latexOk,
      latexSanitized,

      imagesRelCount: Object.keys(mediaRelMap).length,
      imagesOk,
      imagesConverted,
      imagesInjected: 0,

      seenOleRuns: 0,
      seenOle: 0,
      oleInjected: 0,
      ignoredRids: 0,
      sampleRids: [],
      
      // ✅ NEW: sqrt debug info
      sqrtDebug,

      exam: { questions: 0, mcq: 0, tf4: 0, short: 0 },
    };

    const ctx = { latexByRid, imageByRid, debug };

    let inlineHtml = buildInlineHtml(docBuf.toString("utf8"), ctx);
    inlineHtml = formatExamLayout(inlineHtml);
    
    inlineHtml = removeUnsupportedImages(inlineHtml);

    const exam = parseExamFromInlineHtml(inlineHtml);

    if (exam) {
      debug.exam.questions = exam.questions.length;
      for (const q of exam.questions) {
        if (q.type === "mcq") debug.exam.mcq++;
        else if (q.type === "tf4") debug.exam.tf4++;
        else debug.exam.short++;
      }
    }

    return res.json({ ok: true, inlineHtml, exam, debug, mathmlByRid });
  } catch (e) {
    console.error("[CONVERT_DOCX_HTML_FAIL]", e);
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
});
// HẾT ĐOẠN 



const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
