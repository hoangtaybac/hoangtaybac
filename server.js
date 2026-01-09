import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import unzipper from "unzipper";
import { execFileSync, execSync } from "child_process";
import { XMLParser } from "fast-xml-parser";
import { MathMLToLaTeX } from "mathml-to-latex";


// Detect sqrt in MathML (covers common entity forms)
const SQRT_MATHML_RE = /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

/* ================== APP ================== */
const app = express();
app.use(cors());
app.use(express.json({ limit: "25mb" }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/* ================== UTIL ================== */
function safeUnlink(p) {
  try {
    fs.unlinkSync(p);
  } catch {}
}

function safeRmdir(dir) {
  try {
    fs.rmSync(dir, { recursive: true, force: true });
  } catch {}
}

function uniqueTmpPath(baseName = "oleObject.bin") {
  const safe = path.basename(baseName).replace(/[^\w.\-]/g, "_");
  return path.join(
    os.tmpdir(),
    `${Date.now()}_${Math.random().toString(16).slice(2)}_${safe}`
  );
}
async function openDocxZip(docxBuffer) {
  return unzipper.Open.buffer(docxBuffer);
}
async function readZipEntry(zip, p) {
  const f = (zip.files || []).find((x) => x.path === p);
  if (!f) return null;
  return await f.buffer();
}
function unique(arr) {
  return [...new Set(arr || [])].filter(Boolean);
}

/* ================== EMF/WMF -> PNG CONVERSION ================== */
function convertEmfWmfToPng(buffer, ext) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "img-convert-"));
  const inFile = path.join(tmpDir, `input.${ext}`);

  try {
    fs.writeFileSync(inFile, buffer);

    try {
      execSync(
        `soffice --headless --convert-to png "${inFile}" --outdir "${tmpDir}"`,
        { stdio: "ignore", timeout: 30000 }
      );

      const pngFile = fs.readdirSync(tmpDir).find(f => f.endsWith(".png"));
      if (pngFile) {
        return fs.readFileSync(path.join(tmpDir, pngFile));
      }
    } catch (loErr) {
      console.warn("[LIBREOFFICE_CONVERT_WARN]", ext, loErr?.message || String(loErr));
    }

    try {
      const outFile = path.join(tmpDir, "output.png");
      execSync(`convert "${inFile}" "${outFile}"`, { stdio: "ignore", timeout: 30000 });
      if (fs.existsSync(outFile)) {
        return fs.readFileSync(outFile);
      }
    } catch (imErr) {
      console.warn("[IMAGEMAGICK_CONVERT_WARN]", ext, imErr?.message || String(imErr));
    }

    return null;
  } catch (e) {
    console.error("[CONVERT_EMF_WMF_FAIL]", ext, e?.message || String(e));
    return null;
  } finally {
    safeRmdir(tmpDir);
  }
}

/* ================== RUBY OLE(.bin) -> MATHML ================== */
function rubyConvertOleBinToMathML(oleBinBuffer, filenameForTmp) {
  const tmpPath = uniqueTmpPath(path.basename(filenameForTmp || "oleObject.bin"));
  fs.writeFileSync(tmpPath, oleBinBuffer);

  try {
    // Try mt2mml_v2.rb first (handles sqrt properly)
    const v2Script = path.join(process.cwd(), "mt2mml_v2.rb");
    const v1Script = path.join(process.cwd(), "mt2mml.rb");
    const scriptToUse = fs.existsSync(v2Script) ? v2Script : v1Script;
    
    const out = execFileSync(
      "ruby",
      [scriptToUse, tmpPath],
      { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }
    );
    
    // Try to parse as JSON (v2 format)
    let mathml = "";
    try {
      const parsed = JSON.parse(out);
      mathml = parsed.mathml || "";
    } catch {
      // Not JSON, treat as plain MathML (v1 format)
      mathml = (out || "").trim();
    }
    
    if (!mathml || !mathml.startsWith("<")) return "";
    return mathml;
  } catch (e) {
    console.error("[RUBY_FAIL]", {
      file: filenameForTmp,
      msg: e?.message || String(e),
      stderr: e?.stderr ? e.stderr.toString("utf8").slice(0, 1200) : "",
      stdout: e?.stdout ? e.stdout.toString("utf8").slice(0, 1200) : "",
    });
    return "";
  } finally {
    safeUnlink(tmpPath);
  }
}

/* ================== MATHML -> LATEX ================== */

/**
 * ✅ FIX: Pre-process MathML to ensure sqrt elements are properly formatted
 * Some MathType outputs have non-standard sqrt representations
 */
function preprocessMathMLForSqrt(mathml) {
  if (!mathml) return mathml;
  let s = String(mathml);

  // Match sqrt operator inside <mo>...</mo> including common entities
  const moSqrt = String.raw`<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>`;

  // Convert "sqrt operator + following node" into proper <msqrt>...</msqrt>
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mrow>([\s\S]*?)<\/mrow>`, "gi"), "<msqrt>$1</msqrt>");
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mi>([^<]+)<\/mi>`, "gi"), "<msqrt><mi>$1</mi></msqrt>");
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mn>([^<]+)<\/mn>`, "gi"), "<msqrt><mn>$1</mn></msqrt>");
  s = s.replace(new RegExp(moSqrt + String.raw`\s*<mfenced([^>]*)>([\s\S]*?)<\/mfenced>`, "gi"), "<msqrt><mfenced$1>$2</mfenced></msqrt>");

  // IMPORTANT: Do NOT use a "single character after √" rule here; it can break MathML structure.

  return s;
}

/**
 * ✅ FIX: Post-process LaTeX to fix any remaining sqrt issues
 */
function postprocessLatexSqrt(latex) {
  if (!latex) return latex;
  let s = String(latex);
  
  // Some converters output \\surd instead of \\sqrt
  s = s.replace(/\\surd\b/g, '\\sqrt{}');
  
  // Fix: Sometimes √ symbol remains unconverted
  // Pattern: √{content} or √(content) should become \sqrt{content}
  s = s.replace(/√\s*\{([^}]+)\}/g, '\\sqrt{$1}');
  s = s.replace(/√\s*\(([^)]+)\)/g, '\\sqrt{$1}');
  s = s.replace(/√\s*(\d+)/g, '\\sqrt{$1}');
  s = s.replace(/√\s*([a-zA-Z])/g, '\\sqrt{$1}');
  
  // Fix: \sqrt without braces - add braces for single character/number
  s = s.replace(/\\sqrt\s+(\d+)(?![}\d])/g, '\\sqrt{$1}');
  s = s.replace(/\\sqrt\s+([a-zA-Z])(?![}\w])/g, '\\sqrt{$1}');
  
  // Fix: Empty sqrt
  s = s.replace(/\\sqrt\s*\{\s*\}/g, '\\sqrt{\\phantom{x}}');
  
  // Fix: Malformed sqrt with extra spaces
  s = s.replace(/\\sqrt\s+\{/g, '\\sqrt{');
  
  // Fix: nth root - \sqrt[n]{x}
  s = s.replace(/\\root\s*\{([^}]+)\}\s*\\of\s*\{([^}]+)\}/g, '\\sqrt[$1]{$2}');
  s = s.replace(/\\sqrt\s*\[\s*(\d+)\s*\]\s*\{/g, '\\sqrt[$1]{');
  
  return s;
}

/**
 * ✅ FIX: Final LaTeX cleanup - fix Unicode issues, malformed fences, spaced functions
 */
function finalLatexCleanup(latex) {
  if (!latex) return latex;
  let s = String(latex);
  
  // Remove zero-width characters
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '');
  
  // Replace non-breaking spaces with regular spaces
  s = s.replace(/[\u00A0]/g, ' ');
  s = s.replace(/[\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
  
  // Remove control characters
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  
  // Fix: \left( * \right) -> (*)
  s = s.replace(/\\left\s*\(\s*\*\s*\\right\s*\)/g, '(*)');
  s = s.replace(/\\left\s*\(\s*\\star\s*\\right\s*\)/g, '(*)');
  
  // Fix: Malformed \left \right pairs
  s = s.replace(/\\left\s*\(\s*\\right\s*\./g, '(');
  s = s.replace(/\\left\s*\.\s*\\right\s*\)/g, ')');
  s = s.replace(/\\left\s*\(\s*\\right\s*\)/g, '()');
  
  // Fix: Spaced-out functions (l o g -> \log)
  s = s.replace(/\bl\s+o\s+g\b/gi, '\\log');
  s = s.replace(/\bs\s+i\s+n\b/gi, '\\sin');
  s = s.replace(/\bc\s+o\s+s\b/gi, '\\cos');
  s = s.replace(/\bt\s+a\s+n\b/gi, '\\tan');
  s = s.replace(/\bl\s+n\b/gi, '\\ln');
  s = s.replace(/\bl\s+i\s+m\b/gi, '\\lim');
  
  // Fix: log with subscript base
  s = s.replace(/\\log\s*(\d+)\s*_\s*\{\s*\}/g, '\\log_{$1}');
  s = s.replace(/\\log\s+(\d+)\s*\(/g, '\\log_{$1}(');
  s = s.replace(/\\log\s+(\d+)\s*\\left/g, '\\log_{$1}\\left');
  
  // Fix: Empty subscripts/superscripts
  s = s.replace(/_\s*\{\s*\}/g, '');
  s = s.replace(/\^\s*\{\s*\}/g, '');
  
  // Fix: Star symbols
  s = s.replace(/\\star/g, '*');
  s = s.replace(/\\ast/g, '*');
  
  // Clean up multiple spaces
  s = s.replace(/\s{2,}/g, ' ').trim();
  
  return s;
}

/**
 * ✅ FIX: Custom MathML to LaTeX conversion with better sqrt handling
 */
function customMathMLToLatex(mathml) {
  if (!mathml) return "";
  
  // Pre-process to fix sqrt representation
  const preprocessed = preprocessMathMLForSqrt(mathml);
  
  // Try the library first
  let latex = "";
  try {
    latex = MathMLToLaTeX.convert(preprocessed) || "";
  } catch (e) {
    console.error("[MATHML_TO_LATEX_ERROR]", e?.message, "MathML:", mathml.slice(0, 300));
    // Fallback: try manual extraction
    latex = manualMathMLToLatex(preprocessed);
  }
  
  // Post-process to fix any remaining sqrt issues
  latex = postprocessLatexSqrt(latex);
  
  // Log if sqrt was in MathML but not in LaTeX (potential conversion failure)
  if (SQRT_MATHML_RE.test(mathml) 
      && !latex.includes('\\sqrt')) {
    console.warn("[SQRT_LOST] sqrt in MathML but not in LaTeX!");
    console.warn("  MathML:", mathml.slice(0, 500));
    console.warn("  LaTeX:", latex);
    
    // Try manual extraction as fallback
    const manualLatex = manualMathMLToLatex(mathml);
    if (manualLatex.includes('\\sqrt')) {
      console.log("[SQRT_RECOVERED] Manual extraction found sqrt");
      return manualLatex;
    }
  }
  
  return latex.trim();
}

/**
 * ✅ Manual MathML to LaTeX converter as fallback
 * Handles msqrt and mroot specifically
 */
function manualMathMLToLatex(mathml) {
  if (!mathml) return "";
  
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    textNodeName: "#text",
    preserveOrder: false,
  });
  
  let parsed;
  try {
    parsed = parser.parse(mathml);
  } catch (e) {
    console.error("[MANUAL_PARSE_ERROR]", e?.message);
    return "";
  }
  
  function nodeToLatex(node) {
    if (!node) return "";
    if (typeof node === "string") return node;
    if (typeof node === "number") return String(node);
    
    // Handle text node
    if (node["#text"] !== undefined) {
      return String(node["#text"]);
    }
    
    // Handle array of nodes
    if (Array.isArray(node)) {
      return node.map(nodeToLatex).join("");
    }
    
    let result = "";
    
    for (const [tag, content] of Object.entries(node)) {
      if (tag.startsWith("@_")) continue; // Skip attributes
      
      const tagLower = tag.toLowerCase();
      
      switch (tagLower) {
        case "math":
        case "mrow":
        case "mstyle":
        case "mpadded":
        case "mphantom":
          result += nodeToLatex(content);
          break;
          
        case "msqrt":
          result += `\\sqrt{${nodeToLatex(content)}}`;
          break;
          
        case "mroot":
          // mroot has 2 children: base and index
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const index = nodeToLatex(content[1]);
            result += `\\sqrt[${index}]{${base}}`;
          } else {
            result += `\\sqrt{${nodeToLatex(content)}}`;
          }
          break;
          
        case "mfrac":
          if (Array.isArray(content) && content.length >= 2) {
            const num = nodeToLatex(content[0]);
            const den = nodeToLatex(content[1]);
            result += `\\frac{${num}}{${den}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;
          
        case "msup":
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const sup = nodeToLatex(content[1]);
            result += `${base}^{${sup}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;
          
        case "msub":
          if (Array.isArray(content) && content.length >= 2) {
            const base = nodeToLatex(content[0]);
            const sub = nodeToLatex(content[1]);
            result += `${base}_{${sub}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;
          
        case "msubsup":
          if (Array.isArray(content) && content.length >= 3) {
            const base = nodeToLatex(content[0]);
            const sub = nodeToLatex(content[1]);
            const sup = nodeToLatex(content[2]);
            result += `${base}_{${sub}}^{${sup}}`;
          } else {
            result += nodeToLatex(content);
          }
          break;
          
        case "mi":
        case "mn":
        case "mtext":
          result += nodeToLatex(content);
          break;
          
        case "mo":
          const op = nodeToLatex(content);
          // Convert common operators
          const opMap = {
            "√": "\\sqrt",
            "×": "\\times",
            "÷": "\\div",
            "±": "\\pm",
            "∓": "\\mp",
            "≤": "\\leq",
            "≥": "\\geq",
            "≠": "\\neq",
            "≈": "\\approx",
            "∞": "\\infty",
            "→": "\\to",
            "←": "\\leftarrow",
            "⇒": "\\Rightarrow",
            "⇐": "\\Leftarrow",
            "∈": "\\in",
            "∉": "\\notin",
            "⊂": "\\subset",
            "⊃": "\\supset",
            "∪": "\\cup",
            "∩": "\\cap",
            "∀": "\\forall",
            "∃": "\\exists",
            "∂": "\\partial",
            "∇": "\\nabla",
            "∑": "\\sum",
            "∏": "\\prod",
            "∫": "\\int",
            "α": "\\alpha",
            "β": "\\beta",
            "γ": "\\gamma",
            "δ": "\\delta",
            "ε": "\\epsilon",
            "θ": "\\theta",
            "λ": "\\lambda",
            "μ": "\\mu",
            "π": "\\pi",
            "σ": "\\sigma",
            "φ": "\\phi",
            "ω": "\\omega",
          };
          result += opMap[op] || op;
          break;
          
        case "mfenced":
          const open = node["@_open"] || "(";
          const close = node["@_close"] || ")";
          result += `\\left${open}${nodeToLatex(content)}\\right${close}`;
          break;
          
        case "mtable":
          result += `\\begin{matrix}${nodeToLatex(content)}\\end{matrix}`;
          break;
          
        case "mtr":
          result += nodeToLatex(content) + " \\\\ ";
          break;
          
        case "mtd":
          result += nodeToLatex(content) + " & ";
          break;
          
        default:
          result += nodeToLatex(content);
      }
    }
    
    return result;
  }
  
  let latex = nodeToLatex(parsed);
  
  // Clean up
  latex = latex.replace(/\s*&\s*\\\\/g, " \\\\"); // Remove trailing & before \\
  latex = latex.replace(/\s*&\s*$/g, ""); // Remove trailing &
  latex = latex.replace(/\s+/g, " ").trim();
  
  return latex;
}

function mathmlToLatexSafe(mathml) {
  try {
    return customMathMLToLatex(mathml);
  } catch (e) {
    console.error("[MATHML_TO_LATEX_FAIL]", e?.message);
    return "";
  }
}

/**
 * Fix piecewise functions / cases (hệ phương trình, hàm phân đoạn)
 */
function fixPiecewiseFunction(latex) {
  let s = String(latex || "");
  
  // Pattern 1: Fix "(. " -> "(" - MathType's broken parentheses
  s = s.replace(/\(\.\s+/g, "(");
  s = s.replace(/\s+\.\)/g, ")");
  
  // Pattern 2: Fix "[. " -> "[" - MathType's broken brackets  
  s = s.replace(/\[\.\s+/g, "[");
  s = s.replace(/\s+\.\]/g, "]");
  
  // Pattern 3: Find {. pattern (not preceded by \)
  const piecewiseMatch = s.match(/(?<!\\)\{\.\s+/);
  
  if (piecewiseMatch) {
    const startIdx = piecewiseMatch.index;
    const contentStart = startIdx + piecewiseMatch[0].length;
    
    let braceCount = 1;
    let endIdx = contentStart;
    let foundEnd = false;
    
    for (let i = contentStart; i < s.length; i++) {
      const ch = s[i];
      const prevCh = i > 0 ? s[i-1] : "";
      
      if (prevCh === "\\") continue;
      
      if (ch === "{") {
        braceCount++;
      } else if (ch === "}") {
        braceCount--;
        if (braceCount === 0) {
          endIdx = i;
          foundEnd = true;
          break;
        }
      }
    }
    
    if (!foundEnd) {
      endIdx = s.length;
    }
    
    let content = s.slice(contentStart, endIdx).trim();
    content = content.replace(/\s+\.\s*$/, "");
    content = content.replace(/\s+\\\s+(?=\d)/g, " \\\\ ");
    
    const before = s.slice(0, startIdx);
    const after = foundEnd ? s.slice(endIdx + 1) : "";
    
    s = before + `\\begin{cases} ${content} \\end{cases}` + after;
  }
  
  return s;
}

function sanitizeLatexStrict(latex) {
  if (!latex) return latex;

  latex = String(latex).replace(/\s+/g, " ").trim();

  latex = latex
    .replace(/\\left(?!\s*(\(|\[|\\\{|\\langle|\\vert|\\\||\||\.))/g, "")
    .replace(/\\right(?!\s*(\)|\]|\\\}|\\rangle|\\vert|\\\||\||\.))/g, "");

  const tokens = latex.match(/\\left\b|\\right\b/g) || [];
  let bal = 0;
  let broken = false;
  for (const t of tokens) {
    if (t === "\\left") bal++;
    else {
      if (bal === 0) {
        broken = true;
        break;
      }
      bal--;
    }
  }
  if (bal !== 0) broken = true;

  if (broken) latex = latex.replace(/\\left\s*/g, "").replace(/\\right\s*/g, "");
  return latex;
}

function fixSetBracesHard(latex) {
  let s = String(latex || "");

  s = s.replace(/\\underset\s*\{([^}]*)\}\s*\{\s*l\s*i\s*m\s*\}/gi, "\\underset{$1}{\\lim}");
  s = s.replace(/\b(l)\s+(i)\s+(m)\b/gi, "lim");
  s = s.replace(/(^|[^A-Za-z\\])lim([^A-Za-z]|$)/g, "$1\\lim$2");

  s = s.replace(/\\arrow\b/g, "\\rightarrow");
  s = s.replace(/\bxarrow\b/g, "x\\rightarrow");
  s = s.replace(/\\xarrow\b/g, "\\xrightarrow");

  s = s.replace(/\\\{\s*\./g, "\\{");
  s = s.replace(/\.\s*\\\}/g, "\\}");
  s = s.replace(/\\\}\s*\./g, "\\}");

  s = s.replace(/\\mathbb\{([A-Za-z])\\\}/g, "\\mathbb{$1}");
  s = s.replace(/\\mathbb\{([A-Za-z])\}\s*\.\s*\}/g, "\\mathbb{$1}}");

  s = s.replace(/\\backslash\s*{(?!\\)/g, "\\backslash \\{");
  s = s.replace(/\\setminus\s*{(?!\\)/g, "\\setminus \\{");

  if ((s.includes("\\backslash \\{") || s.includes("\\setminus \\{")) && !s.includes("\\}")) {
    s = s.replace(/\}\s*$/g, "").trim() + "\\}";
  }

  s = s.replace(/\\\}\s*([,.;:])/g, "\\}$1");

  s = s.replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, (m, a, b) => {
    const bb = String(b).replace(/(\d)\s+(\d)/g, "$1$2");
    return `\\frac{${a}}{${bb}}`;
  });

  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function restoreArrowAndCoreCommands(latex) {
  let s = String(latex || "");

  s = s.replace(/\s+/g, " ").trim();
  s = s.replace(/\b([A-Za-z])\s+arrow\b/g, "$1 \\to");
  s = s.replace(/\brightarrow\b/g, "\\rightarrow");
  s = s.replace(/\barrow\b/g, "\\rightarrow");
  s = s.replace(/(^|[^A-Za-z\\])to([^A-Za-z]|$)/g, "$1\\to$2");

  return s.replace(/\s+/g, " ").trim();
}

function normalizeLatexCommands(latex) {
  if (!latex) return latex;
  return fixSetBracesHard(String(latex));
}

/* ================== RELS MAP ================== */
function mimeFromExt(p) {
  const ext = (p.split(".").pop() || "").toLowerCase();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "emf") return "image/emf";
  if (ext === "wmf") return "image/wmf";
  return "application/octet-stream";
}

function getExtFromPath(p) {
  return (p.split(".").pop() || "").toLowerCase();
}

function buildRelMaps(relsXmlText) {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const rels = parser.parse(relsXmlText);
  const list = rels?.Relationships?.Relationship || [];
  const arr = Array.isArray(list) ? list : [list];

  const emb = {};
  const media = {};
  for (const r of arr) {
    const id = r?.["@_Id"];
    const target = r?.["@_Target"];
    const targetMode = r?.["@_TargetMode"];
    if (!id || !target) continue;
    if (targetMode && String(targetMode).toLowerCase() === "external") continue;

    const t = String(target).replace(/^\.?\//, "");
    const low = t.toLowerCase();
    if (low.startsWith("embeddings/") && low.endsWith(".bin")) emb[id] = "word/" + t;
    else if (low.startsWith("media/")) media[id] = "word/" + t;
  }
  return { emb, media };
}

/* ================== PRESERVEORDER HELPERS ================== */
function kids(arr, tag) {
  return Array.isArray(arr) ? arr.filter((n) => n && typeof n === "object" && n[tag]) : [];
}
function findAllRidsDeep(x, out = []) {
  const re = /^rId\d+$/;
  if (!x) return out;

  if (typeof x === "string") {
    const s = x.trim();
    if (re.test(s)) out.push(s);
    return out;
  }
  if (Array.isArray(x)) {
    for (const it of x) findAllRidsDeep(it, out);
    return out;
  }
  if (typeof x === "object") {
    for (const v of Object.values(x)) findAllRidsDeep(v, out);
    return out;
  }
  return out;
}
function findImageEmbedRidsDeep(x, out = []) {
  if (!x) return out;
  if (Array.isArray(x)) {
    for (const it of x) findImageEmbedRidsDeep(it, out);
    return out;
  }
  if (typeof x === "object") {
    for (const [k, v] of Object.entries(x)) {
      if ((k === "@_r:embed" || k === "@_r:id") && typeof v === "string" && v.startsWith("rId")) out.push(v);
      findImageEmbedRidsDeep(v, out);
    }
  }
  return out;
}
function runHasOleLike(rNode) {
  try {
    const s = JSON.stringify(rNode);
    return s.includes("o:OLEObject") || s.includes("w:object") || s.includes("w:oleObject");
  } catch {
    return false;
  }
}
function runIsUnderlined(rNode) {
  try {
    const s = JSON.stringify(rNode);
    if (!s.includes("w:u")) return false;
    if (s.toLowerCase().includes("none")) return false;
    return true;
  } catch {
    return false;
  }
}

/* ================== TEXT EXTRACTION ================== */
function getTextFromPreserveWrap(tagWrap, tagName) {
  const v = tagWrap?.[tagName];
  if (!v) return "";
  if (Array.isArray(v)) return v.map((x) => x?.["#text"] || "").join("");
  if (typeof v === "object") return v?.["#text"] || "";
  return "";
}
function collectTextFromRun(rNode) {
  let s = "";
  for (const tWrap of kids(rNode, "w:t")) s += getTextFromPreserveWrap(tWrap, "w:t");
  for (const tWrap of kids(rNode, "w:instrText")) s += getTextFromPreserveWrap(tWrap, "w:instrText");
  for (const tWrap of kids(rNode, "w:delText")) s += getTextFromPreserveWrap(tWrap, "w:delText");
  if (kids(rNode, "w:tab").length) s += "\t";
  if (kids(rNode, "w:br").length) s += "\n";
  return s;
}
function escapeTextToHtml(text) {
  if (!text) return "";
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\t", "&emsp;")
    .replaceAll("\n", "<br/>");
}

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
          childRids.forEach(rid => processedInLoop.add(rid));
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
            childRids.forEach(rid => {
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
  let last = 0, m;
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

function formatExamLayout(html) {
  let result = html;
  
  result = result.replace(/\s+/g, " ");
  result = result.replace(/PHẦN(\d)/gi, "PHẦN $1");
  
  result = result.replace(
    /(^|<br\/>)\s*(PHẦN\s+\d+\.(?:(?!<br\/>\s*Câu\s+\d).)*)/g,
    '$1<br/><div class="section-header"><strong>$2</strong></div>'
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

function formatAbcdOutsideHeaders(text) {
  const headerRegex = /(<div class="section-header">[\s\S]*?<\/div>)/g;
  const segments = text.split(headerRegex);
  
  return segments.map(seg => {
    if (seg.startsWith('<div class="section-header">')) {
      return seg;
    }
    
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
  }).join('');
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
  const re = /(Lời(?:\s*<[^>]*>)*\s*giải|Giải(?:\s*<[^>]*>)*\s*chi\s*tiết|Hướng(?:\s*<[^>]*>)*\s*dẫn(?:\s*<[^>]*>)*\s*giải)/i;
  const sub = s.slice(fromIndex);
  const m = re.exec(sub);
  if (!m) return -1;
  return fromIndex + m.index;
}

function splitSolutionSections(tailHtml) {
  let s = String(tailHtml || "").trim();
  if (!s) return { solutionHtml: "", detailHtml: "" };

  const reCT = /(Giải(?:\s*<[^>]*>)*\s*chi\s*tiết)/i;
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

/* ================== START ================== */
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
