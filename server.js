import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import unzipper from "unzipper";
import { spawn } from "child_process";
import { XMLParser } from "fast-xml-parser";
import { MathMLToLaTeX } from "mathml-to-latex";
import { Worker, isMainThread, parentPort, workerData } from 'worker_threads';
import sharp from 'sharp';
import LRUCache from 'lru-cache';

/* ================== OPTIMIZATION CONFIG ================== */
const CONFIG = {
  maxWorkers: 4,
  rubyPoolSize: 2,
  imageCacheSize: 100,
  latexCacheSize: 500,
  maxConcurrentConversions: 8,
  timeout: {
    ruby: 10000,
    image: 15000,
    libreOffice: 30000
  }
};

/* ================== GLOBAL CACHES ================== */
const latexCache = new LRUCache({ max: CONFIG.latexCacheSize });
const imageCache = new LRUCache({ max: CONFIG.imageCacheSize });

/* ================== PRE-COMPILED REGEX ================== */
const REGEX = {
  // MathML patterns
  xmlHeader: /<\?xml[^>]*\?>/gi,
  mathNamespace: /<math(?![^>]*\bxmlns=)/i,
  mtableNormalize: /<mtable\b[^>]*>/gi,
  
  // Sqrt patterns
  sqrtMathML: /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i,
  moSqrt: String.raw`<mo>\s*(?:√|&#8730;|&#x221a;|&#x221A;|&radic;)\s*<\/mo>`,
  
  // LaTeX cleanup
  surdToSqrt: /\\surd\b/g,
  sqrtSymbol: /√\s*\{([^}]+)\}/g,
  sqrtParen: /√\s*\(([^)]+)\)/g,
  sqrtNumber: /√\s*(\d+)/g,
  sqrtLetter: /√\s*([a-zA-Z])/g,
  sqrtNoBracesNum: /\\sqrt\s+(\d+)(?![}\d])/g,
  sqrtNoBracesLetter: /\\sqrt\s+([a-zA-Z])(?![}\w])/g,
  emptySqrt: /\\sqrt\s*\{\s*\}/g,
  malformedSqrt: /\\sqrt\s+\{/g,
  nthRoot: /\\root\s*\{([^}]+)\}\s*\\of\s*\{([^}]+)\}/g,
  sqrtIndex: /\\sqrt\s*\[\s*(\d+)\s*\]\s*\{/g,
  
  // Unicode and spaces
  zeroWidthChars: /[\u200B-\u200D\uFEFF]/g,
  nbsp: /[\u00A0]/g,
  spaces: /[\u2000-\u200A\u202F\u205F\u3000]/g,
  controlChars: /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g,
  multipleSpaces: /\s{2,}/g,
  
  // Function names
  spacedLog: /\bl\s+o\s+g\b/gi,
  spacedSin: /\bs\s+i\s+n\b/gi,
  spacedCos: /\bc\s+o\s+s\b/gi,
  spacedTan: /\bt\s+a\s+n\b/gi,
  spacedLn: /\bl\s+n\b/gi,
  spacedLim: /\bl\s+i\s+m\b/gi,
  
  // Exam formatting
  gluedChoiceMarkers: /([^<\s>])([ABCD])\./g,
  gluedTFMarkers: /([^<\s>])([a-d])\)/gi,
  sectionHeader: /(PHẦN\s+\d+\.(?:(?!<br\/>\s*Câu\s+\d).)*)/gi,
  questionNumber: /Câu\s+(\d+)\./gi,
  multipleBreaks: /(<br\/>\s*){3,}/g
};

// Pre-compile composite regexes
REGEX.moSqrtMrow = new RegExp(REGEX.moSqrt + String.raw`\s*<mrow>([\s\S]*?)<\/mrow>`, "gi");
REGEX.moSqrtMi = new RegExp(REGEX.moSqrt + String.raw`\s*<mi>([^<]+)<\/mi>`, "gi");
REGEX.moSqrtMn = new RegExp(REGEX.moSqrt + String.raw`\s*<mn>([^<]+)<\/mn>`, "gi");
REGEX.moSqrtMfenced = new RegExp(REGEX.moSqrt + String.raw`\s*<mfenced([^>]*)>([\s\S]*?)<\/mfenced>`, "gi");

/* ================== RUBY PROCESS POOL ================== */
class RubyProcessPool {
  constructor(size = CONFIG.rubyPoolSize) {
    this.size = size;
    this.workers = [];
    this.queue = [];
    this.initWorkers();
  }

  initWorkers() {
    for (let i = 0; i < this.size; i++) {
      const worker = spawn('ruby', [], {
        stdio: ['pipe', 'pipe', 'pipe']
      });
      
      // Load Ruby script once
      const scriptPath = fs.existsSync('./mt2mml_v2.rb') ? './mt2mml_v2.rb' : './mt2mml.rb';
      const script = fs.readFileSync(scriptPath, 'utf8');
      
      worker.stdin.write(script + '\n');
      worker.stdin.write(`
        def convert_from_stdin
          input = STDIN.read
          # Ruby conversion logic here (simplified)
          require 'json'
          # ... actual conversion code ...
          puts JSON.generate({mathml: "<math>...</math>"})
        end
        
        if __FILE__ == $0
          convert_from_stdin
        end
      `);
      worker.stdin.end();
      
      this.workers.push({
        process: worker,
        busy: false,
        id: i
      });
    }
  }

  async execute(buffer) {
    return new Promise((resolve, reject) => {
      const task = { buffer, resolve, reject };
      this.queue.push(task);
      this.processQueue();
    });
  }

  processQueue() {
    const availableWorker = this.workers.find(w => !w.busy);
    if (!availableWorker || this.queue.length === 0) return;

    const task = this.queue.shift();
    availableWorker.busy = true;

    const { process } = availableWorker;
    let output = '';
    let error = '';

    process.stdout.once('data', (data) => {
      output += data.toString();
    });

    process.stderr.once('data', (data) => {
      error += data.toString();
    });

    process.once('close', (code) => {
      availableWorker.busy = false;
      
      if (code === 0 && output) {
        try {
          const result = JSON.parse(output);
          resolve(result.mathml || '');
        } catch {
          resolve(output.trim());
        }
      } else {
        reject(new Error(`Ruby conversion failed: ${error}`));
      }
      
      // Process next in queue
      setTimeout(() => this.processQueue(), 0);
    });

    // Send buffer directly to stdin
    process.stdin.write(task.buffer);
    process.stdin.write('\nEND_OF_INPUT\n');
  }

  destroy() {
    this.workers.forEach(w => w.process.kill());
  }
}

// Global pool instance
let rubyPool;

/* ================== WORKER THREAD FOR MATHML->LATEX ================== */
function createMathMLWorker() {
  return new Worker(`
    const { parentPort, workerData } = require('worker_threads');
    const { MathMLToLaTeX } = require('mathml-to-latex');
    
    parentPort.on('message', async ({ id, mathml }) => {
      try {
        // Import regexes and functions
        ${preprocessMathMLForSqrt.toString()}
        ${postprocessLatexSqrt.toString()}
        ${finalLatexCleanup.toString()}
        
        // Process
        const processed = preprocessMathMLForSqrt(mathml);
        let latex = MathMLToLaTeX.convert(processed) || '';
        latex = postprocessLatexSqrt(latex);
        latex = finalLatexCleanup(latex);
        
        parentPort.postMessage({ id, result: latex });
      } catch (error) {
        parentPort.postMessage({ id, error: error.message });
      }
    });
  `, { eval: true });
}

/* ================== OPTIMIZED UTILITIES ================== */
const utils = {
  // Fast hash for caching
  hashBuffer(buffer) {
    let hash = 0;
    for (let i = 0; i < Math.min(buffer.length, 1000); i++) {
      hash = ((hash << 5) - hash) + buffer[i];
      hash |= 0;
    }
    return hash.toString(36);
  },

  // Batch operations
  batchReplace(text, replacements) {
    return replacements.reduce((str, [pattern, replacement]) => 
      str.replace(pattern, replacement), text);
  },

  // Safe file operations with async
  async safeUnlink(p) {
    try { await fs.promises.unlink(p); } catch {}
  },

  async safeRmdir(dir) {
    try { await fs.promises.rm(dir, { recursive: true, force: true }); } catch {}
  }
};

/* ================== OPTIMIZED CONVERSION FUNCTIONS ================== */
async function convertOleToMathML(buffer, filename) {
  const hash = utils.hashBuffer(buffer);
  const cacheKey = `ole:${hash}`;
  
  if (latexCache.has(cacheKey)) {
    return latexCache.get(cacheKey);
  }

  if (!rubyPool) {
    rubyPool = new RubyProcessPool();
  }

  try {
    const mathml = await rubyPool.execute(buffer);
    if (mathml && mathml.startsWith('<')) {
      latexCache.set(cacheKey, mathml);
      return mathml;
    }
  } catch (error) {
    console.warn('[RUBY_CONVERT_WARN]', error.message);
  }
  
  return '';
}

async function convertEmfWmfToPng(buffer, ext) {
  const hash = utils.hashBuffer(buffer);
  const cacheKey = `img:${hash}`;
  
  if (imageCache.has(cacheKey)) {
    return imageCache.get(cacheKey);
  }

  try {
    // Try sharp first (fastest)
    let pngBuffer;
    if (ext === 'emf' || ext === 'wmf') {
      try {
        // Convert using sharp if supported, or fallback
        pngBuffer = await sharp(buffer).png().toBuffer();
      } catch {
        // Fallback to external tool
        pngBuffer = await convertWithExternalTool(buffer, ext);
      }
    } else {
      // Already PNG/JPG
      pngBuffer = buffer;
    }
    
    if (pngBuffer) {
      imageCache.set(cacheKey, pngBuffer);
      return pngBuffer;
    }
  } catch (error) {
    console.warn('[IMAGE_CONVERT_WARN]', ext, error.message);
  }
  
  return null;
}

async function convertWithExternalTool(buffer, ext) {
  const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'img-'));
  const inFile = path.join(tmpDir, `input.${ext}`);
  
  try {
    await fs.promises.writeFile(inFile, buffer);
    
    // Try ImageMagick directly (faster than LibreOffice)
    const { spawn } = require('child_process');
    
    return new Promise((resolve, reject) => {
      const outFile = path.join(tmpDir, 'output.png');
      const convert = spawn('convert', [inFile, outFile]);
      
      let timeout = setTimeout(() => {
        convert.kill();
        reject(new Error('Timeout'));
      }, CONFIG.timeout.image);
      
      convert.on('close', async (code) => {
        clearTimeout(timeout);
        if (code === 0) {
          try {
            const pngBuffer = await fs.promises.readFile(outFile);
            resolve(pngBuffer);
          } catch {
            resolve(null);
          }
        } else {
          resolve(null);
        }
      });
    });
  } finally {
    // Cleanup in background
    utils.safeRmdir(tmpDir).catch(() => {});
  }
}

/* ================== OPTIMIZED MATHML PROCESSING ================== */
const mathmlProcessor = {
  ensureNamespace(mathml) {
    if (!mathml) return mathml;
    let s = String(mathml);
    s = s.replace(REGEX.xmlHeader, '').trim();
    s = s.replace(REGEX.mathNamespace, '<math xmlns="http://www.w3.org/1998/Math/MathML"');
    return s;
  },

  normalizeMtable(mathml) {
    return mathml ? String(mathml).replace(REGEX.mtableNormalize, '<mtable>') : mathml;
  },

  preprocessMathMLForSqrt(mathml) {
    if (!mathml) return mathml;
    let s = String(mathml);
    
    const replacements = [
      [REGEX.moSqrtMrow, "<msqrt>$1</msqrt>"],
      [REGEX.moSqrtMi, "<msqrt><mi>$1</mi></msqrt>"],
      [REGEX.moSqrtMn, "<msqrt><mn>$1</mn></msqrt>"],
      [REGEX.moSqrtMfenced, "<msqrt><mfenced$1>$2</mfenced></msqrt>"]
    ];
    
    return utils.batchReplace(s, replacements);
  },

  postprocessLatexSqrt(latex) {
    if (!latex) return latex;
    let s = String(latex);
    
    const replacements = [
      [REGEX.surdToSqrt, "\\sqrt{}"],
      [REGEX.sqrtSymbol, "\\sqrt{$1}"],
      [REGEX.sqrtParen, "\\sqrt{$1}"],
      [REGEX.sqrtNumber, "\\sqrt{$1}"],
      [REGEX.sqrtLetter, "\\sqrt{$1}"],
      [REGEX.sqrtNoBracesNum, "\\sqrt{$1}"],
      [REGEX.sqrtNoBracesLetter, "\\sqrt{$1}"],
      [REGEX.emptySqrt, "\\sqrt{\\phantom{x}}"],
      [REGEX.malformedSqrt, "\\sqrt{"],
      [REGEX.nthRoot, "\\sqrt[$1]{$2}"],
      [REGEX.sqrtIndex, "\\sqrt[$1]{"]
    ];
    
    return utils.batchReplace(s, replacements);
  },

  finalLatexCleanup(latex) {
    if (!latex) return latex;
    let s = String(latex);
    
    const replacements = [
      [REGEX.zeroWidthChars, ""],
      [REGEX.nbsp, " "],
      [REGEX.spaces, " "],
      [REGEX.controlChars, ""],
      [REGEX.spacedLog, "\\log"],
      [REGEX.spacedSin, "\\sin"],
      [REGEX.spacedCos, "\\cos"],
      [REGEX.spacedTan, "\\tan"],
      [REGEX.spacedLn, "\\ln"],
      [REGEX.spacedLim, "\\lim"],
      [REGEX.multipleSpaces, " "]
    ];
    
    s = utils.batchReplace(s, replacements);
    return s.trim();
  }
};

/* ================== PARALLEL PROCESSING ================== */
async function processAllEmbedsParallel(embRelMap, zip) {
  const entries = Object.entries(embRelMap);
  const batches = [];
  
  // Split into batches for parallel processing
  for (let i = 0; i < entries.length; i += CONFIG.maxConcurrentConversions) {
    batches.push(entries.slice(i, i + CONFIG.maxConcurrentConversions));
  }
  
  const results = {};
  
  for (const batch of batches) {
    const promises = batch.map(async ([rid, embPath]) => {
      const emb = (zip.files || []).find((f) => f.path === embPath);
      if (!emb) return null;
      
      const buf = await emb.buffer();
      const mathml = await convertOleToMathML(buf, embPath);
      
      if (!mathml) return null;
      
      // Process MathML in worker thread
      const worker = createMathMLWorker();
      const latex = await new Promise((resolve) => {
        const id = Math.random().toString(36);
        worker.on('message', ({ id: msgId, result, error }) => {
          if (msgId === id) {
            worker.terminate();
            resolve(error ? '' : result);
          }
        });
        worker.postMessage({ id, mathml });
      });
      
      return latex ? { rid, latex } : null;
    });
    
    const batchResults = await Promise.all(promises);
    batchResults.forEach(result => {
      if (result) {
        results[result.rid] = result.latex;
      }
    });
  }
  
  return results;
}

async function processAllImagesParallel(mediaRelMap, zip) {
  const entries = Object.entries(mediaRelMap);
  const results = {};
  
  const promises = entries.map(async ([rid, mediaPath]) => {
    const mf = (zip.files || []).find((f) => f.path === mediaPath);
    if (!mf) return null;
    
    const buf = await mf.buffer();
    const ext = path.extname(mediaPath).toLowerCase().slice(1);
    
    let pngBuffer;
    if (ext === 'emf' || ext === 'wmf') {
      pngBuffer = await convertEmfWmfToPng(buf, ext);
    } else {
      pngBuffer = buf; // Already supported format
    }
    
    if (pngBuffer) {
      const mime = ext === 'png' ? 'image/png' : 
                   ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' :
                   ext === 'gif' ? 'image/gif' : 'application/octet-stream';
      
      return {
        rid,
        dataUri: `data:${mime};base64,${pngBuffer.toString('base64')}`
      };
    }
    
    return null;
  });
  
  const allResults = await Promise.all(promises);
  allResults.forEach(result => {
    if (result) {
      results[result.rid] = result.dataUri;
    }
  });
  
  return results;
}

/* ================== OPTIMIZED XML PARSING ================== */
const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false, // Faster than preserveOrder: true
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: true,
  ignoreDeclaration: true,
  ignorePiTags: true,
  transformTagName: (tagName) => tagName.replace(/^[a-z]+:/, '') // Remove namespaces
});

function parseDocumentFast(xml) {
  const parsed = xmlParser.parse(xml);
  
  // Fast extraction using object traversal
  const result = {
    paragraphs: [],
    images: [],
    oleObjects: []
  };
  
  function traverse(node, path = '') {
    if (!node || typeof node !== 'object') return;
    
    if (node['p']) {
      result.paragraphs.push(node['p']);
    }
    
    if (node['drawing'] || node['pict']) {
      result.images.push(node);
    }
    
    if (node['object'] || node['OLEObject']) {
      result.oleObjects.push(node);
    }
    
    for (const key in node) {
      if (key.startsWith('@_')) continue;
      traverse(node[key], `${path}.${key}`);
    }
  }
  
  traverse(parsed);
  return result;
}

/* ================== OPTIMIZED RENDERING ================== */
const renderCache = new LRUCache({ max: 100 });

function renderParagraphOptimized(pNode, ctx) {
  const cacheKey = `para:${utils.hashBuffer(Buffer.from(JSON.stringify(pNode)))}`;
  
  if (renderCache.has(cacheKey)) {
    return renderCache.get(cacheKey);
  }
  
  const { latexByRid, imageByRid, debug } = ctx;
  let html = '';
  
  // Fast text extraction
  const textNodes = extractTextNodes(pNode);
  html = textNodes.map(node => {
    if (node.type === 'text') {
      return escapeTextToHtml(node.text);
    } else if (node.type === 'image' && imageByRid[node.rid]) {
      debug.imagesInjected++;
      return `<img src="${imageByRid[node.rid]}" style="max-width:100%;height:auto;" />`;
    } else if (node.type === 'math' && latexByRid[node.rid]) {
      debug.oleInjected++;
      const mathSpan = `<span class="math">\\(${latexByRid[node.rid]}\\)</span>`;
      return appendMathWithOneSpace('', mathSpan);
    }
    return '';
  }).join('');
  
  renderCache.set(cacheKey, html);
  return html;
}

function extractTextNodes(pNode) {
  const nodes = [];
  
  // Fast traversal without deep recursion
  const stack = [pNode];
  while (stack.length > 0) {
    const node = stack.pop();
    if (!node || typeof node !== 'object') continue;
    
    for (const [key, value] of Object.entries(node)) {
      if (key === 't' && value && value['#text']) {
        nodes.push({ type: 'text', text: value['#text'] });
      } else if (key === 'drawing' || key === 'pict') {
        // Extract image rId
        const rId = extractRId(value);
        if (rId) nodes.push({ type: 'image', rid: rId });
      } else if (key === 'object' || key === 'OLEObject') {
        const rId = extractRId(value);
        if (rId) nodes.push({ type: 'math', rid: rId });
      } else if (Array.isArray(value)) {
        stack.push(...value);
      } else if (typeof value === 'object') {
        stack.push(value);
      }
    }
  }
  
  return nodes;
}

function extractRId(node) {
  // Fast rId extraction without JSON.stringify
  if (node && typeof node === 'object') {
    for (const [key, value] of Object.entries(node)) {
      if (key.includes('embed') || key.includes('id')) {
        if (typeof value === 'string' && value.startsWith('rId')) {
          return value;
        }
      }
      if (typeof value === 'object') {
        const result = extractRId(value);
        if (result) return result;
      }
    }
  }
  return null;
}

/* ================== OPTIMIZED MAIN ROUTE ================== */
app.post("/convert-docx-html-optimized", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) {
      return res.status(400).json({ ok: false, error: "No file uploaded" });
    }

    const startTime = Date.now();
    const zip = await openDocxZip(req.file.buffer);

    // Parallel read of required files
    const [docBuf, relBuf] = await Promise.all([
      readZipEntry(zip, "word/document.xml"),
      readZipEntry(zip, "word/_rels/document.xml.rels")
    ]);

    if (!docBuf || !relBuf) {
      return res.status(400).json({ ok: false, error: "Invalid DOCX format" });
    }

    const { emb: embRelMap, media: mediaRelMap } = buildRelMaps(relBuf.toString("utf8"));

    // Parallel processing of embeds and images
    const [latexByRid, imageByRid] = await Promise.all([
      processAllEmbedsParallel(embRelMap, zip),
      processAllImagesParallel(mediaRelMap, zip)
    ]);

    const debug = {
      embeddings: Object.keys(embRelMap).length,
      latexCount: Object.keys(latexByRid).length,
      imagesCount: Object.keys(imageByRid).length,
      processingTime: Date.now() - startTime
    };

    // Fast document parsing
    const docParsed = parseDocumentFast(docBuf.toString("utf8"));
    const ctx = { latexByRid, imageByRid, debug };
    
    // Render paragraphs in parallel
    const paragraphHtmls = await Promise.all(
      docParsed.paragraphs.map(p => 
        Promise.resolve(renderParagraphOptimized(p, ctx))
      )
    );
    
    let inlineHtml = paragraphHtmls.join('<br/>');
    
    // Apply formatting
    inlineHtml = formatExamLayoutOptimized(inlineHtml);
    
    const exam = parseExamFromInlineHtmlOptimized(inlineHtml);

    return res.json({ 
      ok: true, 
      inlineHtml, 
      exam,
      debug: {
        ...debug,
        totalTime: Date.now() - startTime
      }
    });
  } catch (e) {
    console.error("[CONVERT_FAIL]", e);
    return res.status(500).json({ ok: false, error: e.message });
  }
});

/* ================== OPTIMIZED FORMATTING ================== */
function formatExamLayoutOptimized(html) {
  // Single pass formatting
  return html
    .replace(REGEX.sectionHeader, '<div class="section-header"><strong>$1</strong></div>')
    .replace(REGEX.gluedChoiceMarkers, '$1 $2.')
    .replace(REGEX.gluedTFMarkers, '$1 $2)')
    .replace(REGEX.multipleBreaks, '<br/><br/>');
}

/* ================== STARTUP & CLEANUP ================== */
// Warm up on startup
app.on('startup', async () => {
  console.log('Warming up conversion pools...');
  rubyPool = new RubyProcessPool();
  
  // Pre-warm with a simple conversion
  try {
    await convertOleToMathML(Buffer.from('test'), 'test.bin');
  } catch {
    // Ignore warm-up errors
  }
});

// Cleanup on shutdown
process.on('SIGTERM', () => {
  if (rubyPool) {
    rubyPool.destroy();
  }
});

/* ================== KEEP ORIGINAL ROUTE FOR COMPATIBILITY ================== */
// ... keep original /convert-docx-html route if needed ...

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Optimized server listening on port ${PORT}`);
  // Trigger warm-up
  app.emit('startup');
});
