<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Test MathType Converter</title>

  <!-- MathJax (TeX + MathML) -->
  <script>
    window.MathJax = {
      startup: { typeset: false },
      tex: {
        inlineMath: [["\\(", "\\)"]],
        displayMath: [["\\[", "\\]"]],
        processEscapes: true
      },
      options: {
        skipHtmlTags: ["script","noscript","style","textarea","pre","code"]
      }
    };
  </script>
  <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>

  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{
      font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;
      background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
      min-height:100vh;padding:20px
    }
    .container{
      max-width:1200px;margin:0 auto;background:white;border-radius:20px;
      box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden
    }
    .header{
      background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
      color:white;padding:30px;text-align:center
    }
    .header h1{font-size:2em;margin-bottom:10px}
    .header p{opacity:.9;font-size:1.1em}
    .upload-section{padding:40px;border-bottom:2px solid #f0f0f0}
    .file-input-wrapper{position:relative;display:inline-block;width:100%;margin-bottom:20px}
    .file-input-wrapper input[type="file"]{position:absolute;opacity:0;width:100%;height:100%;cursor:pointer}
    .file-input-label{
      display:block;padding:30px;background:#f8f9fa;border:3px dashed #667eea;border-radius:15px;
      text-align:center;cursor:pointer;transition:all .3s ease
    }
    .file-input-label:hover{background:#e9ecef;border-color:#764ba2}
    .file-input-label.has-file{background:#d4edda;border-color:#28a745}
    .upload-btn{
      width:100%;padding:15px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
      color:white;border:none;border-radius:10px;font-size:1.1em;font-weight:bold;cursor:pointer;
      transition:transform .2s
    }
    .upload-btn:hover:not(:disabled){
      transform:translateY(-2px);box-shadow:0 5px 15px rgba(102,126,234,.4)
    }
    .upload-btn:disabled{background:#ccc;cursor:not-allowed}
    .loading{text-align:center;padding:40px;display:none}
    .spinner{
      border:4px solid #f3f3f3;border-top:4px solid #667eea;border-radius:50%;
      width:50px;height:50px;animation:spin 1s linear infinite;margin:0 auto 20px
    }
    @keyframes spin{0%{transform:rotate(0)}100%{transform:rotate(360deg)}}
    .results{display:none;padding:40px}
    .tab-buttons{
      display:flex;gap:10px;margin-bottom:20px;border-bottom:2px solid #e9ecef;flex-wrap:wrap
    }
    .tab-btn{
      padding:12px 24px;background:none;border:none;border-bottom:3px solid transparent;
      cursor:pointer;font-size:1em;font-weight:600;color:#666;transition:all .3s
    }
    .tab-btn.active{color:#667eea;border-bottom-color:#667eea}
    .tab-content{display:none;animation:fadeIn .3s}
    .tab-content.active{display:block}
    @keyframes fadeIn{from{opacity:0}to{opacity:1}}
    .debug-info{background:#f8f9fa;padding:20px;border-radius:10px;margin-bottom:20px}
    .debug-info h3{color:#667eea;margin-bottom:15px}
    .debug-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px}
    .debug-item{background:white;padding:15px;border-radius:8px;border-left:4px solid #667eea}
    .debug-item label{display:block;font-size:.85em;color:#666;margin-bottom:5px}
    .debug-item .value{font-size:1.5em;font-weight:bold;color:#333}
    .html-preview{background:white;padding:30px;border:1px solid #dee2e6;border-radius:10px;line-height:1.8}
    .html-preview img{max-width:100%;height:auto;margin:10px 0}

    .question-card{
      background:white;border:2px solid #e9ecef;border-radius:10px;padding:25px;margin-bottom:20px;transition:all .3s
    }
    .question-card:hover{border-color:#667eea;box-shadow:0 5px 15px rgba(102,126,234,.1)}
    .question-header{
      display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;padding-bottom:15px;
      border-bottom:2px solid #f0f0f0
    }
    .question-number{font-size:1.2em;font-weight:bold;color:#667eea}
    .question-type{padding:5px 15px;border-radius:20px;font-size:.85em;font-weight:bold}
    .type-mcq{background:#d4edda;color:#155724}
    .type-tf4{background:#cce5ff;color:#004085}
    .type-short{background:#fff3cd;color:#856404}
    .question-stem{margin-bottom:20px;font-size:1.1em;line-height:1.6}

    .choices{display:grid;gap:10px}
    .choice{padding:15px;background:#f8f9fa;border-radius:8px;border-left:4px solid transparent;transition:all .2s}
    .choice.correct{background:#d4edda;border-left-color:#28a745}
    .choice-label{font-weight:bold;color:#667eea;margin-right:10px}

    .tf-pill{
      display:inline-block;
      padding:2px 10px;
      border-radius:999px;
      font-size:.85em;
      font-weight:700;
      margin-left:8px;
      vertical-align:middle;
    }
    .pill-true{background:#d4edda;color:#155724;border:1px solid #a9dfb4}
    .pill-false{background:#f8d7da;color:#721c24;border:1px solid #f1aeb5}
    .pill-unknown{background:#e2e3e5;color:#41464b;border:1px solid #d3d6d8}

    .error{background:#f8d7da;color:#721c24;padding:20px;border-radius:10px;margin:20px}
    .json-view{
      background:#282c34;color:#abb2bf;padding:20px;border-radius:10px;overflow-x:auto;
      font-family:'Courier New',monospace;font-size:.9em;line-height:1.6
    }
    .stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:20px;margin-bottom:30px}
    .stat-card{
      background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
      color:white;padding:25px;border-radius:15px;text-align:center
    }
    .stat-value{font-size:3em;font-weight:bold;margin-bottom:10px}
    .stat-label{font-size:1.1em;opacity:.9}

    .solution-toggle-btn{
      margin-top:15px;padding:10px 20px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
      color:white;border:none;border-radius:8px;font-size:1em;font-weight:600;cursor:pointer;
      transition:all .3s;display:block;width:100%
    }
    .solution-toggle-btn:hover{transform:translateY(-2px);box-shadow:0 5px 15px rgba(102,126,234,.4)}
    .solution-toggle-btn.active{background:linear-gradient(135deg,#764ba2 0%,#667eea 100%)}
    .solution-content{
      margin-top:15px;padding:20px;background:#f8f9fa;border-radius:10px;border-left:4px solid #667eea
    }

    img.docx-img{max-width:100%;height:auto;border-radius:8px}
    img.inline-img{display:inline-block;vertical-align:middle;max-height:1.6em;margin:0 2px}
  </style>
</head>

<body>
<div class="container">
  <div class="header">
    <h1>🧪 MathType Converter Test</h1>
    <p>Upload file Word để test parsing đề thi</p>
    <p style="font-size:.9em;margin-top:10px;">
      Server: <strong>hoangtaybac-production.up.railway.app</strong>
    </p>
  </div>

  <div class="upload-section">
    <div class="file-input-wrapper">
      <input type="file" id="fileInput" accept=".docx" onchange="handleFileSelect(event)">
      <label class="file-input-label" id="fileLabel">
        <div style="font-size:3em;margin-bottom:10px;">📄</div>
        <div style="font-size:1.2em;font-weight:bold;margin-bottom:5px;">Chọn file Word (.docx)</div>
        <div style="color:#666;">Hoặc kéo thả file vào đây</div>
      </label>
    </div>

    <button class="upload-btn" id="uploadBtn" onclick="uploadFile()" disabled>
      🚀 Upload và Phân Tích
    </button>
  </div>

  <div class="loading" id="loading">
    <div class="spinner"></div>
    <h3>Đang xử lý file...</h3>
    <p>Vui lòng đợi trong giây lát</p>
  </div>

  <div class="results" id="results">
    <div class="stats-grid" id="statsGrid"></div>

    <div class="tab-buttons">
      <button class="tab-btn active" onclick="switchTab('overview', this)">📊 Tổng Quan</button>
      <button class="tab-btn" onclick="switchTab('questions', this)">📝 Câu Hỏi</button>
      <button class="tab-btn" onclick="switchTab('latex', this)">🔬 LaTeX Debug</button>
      <button class="tab-btn" onclick="switchTab('html', this)">🌐 HTML Preview</button>
      <button class="tab-btn" onclick="switchTab('debug', this)">🔧 Debug Info</button>
      <button class="tab-btn" onclick="switchTab('json', this)">💻 Raw JSON</button>
    </div>

    <div id="tab-overview" class="tab-content active"></div>
    <div id="tab-questions" class="tab-content"></div>
    <div id="tab-latex" class="tab-content"></div>
    <div id="tab-html" class="tab-content"></div>
    <div id="tab-debug" class="tab-content"></div>
    <div id="tab-json" class="tab-content"></div>
  </div>
</div>

<script>
  const API_URL = 'https://hoangtaybac-production.up.railway.app';
  let selectedFile = null;

  // from server
  let latexMap = {};
  let imgMap = {};
  let rawText = "";
  let questions = [];

  function escapeHtml(s="") {
    return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  }

  function renderInline(text) {
    if (!text) return "";
    let html = escapeHtml(text).replace(/\t/g,"&emsp;").replace(/\n/g,"<br>");

    // img tokens
    html = html.replace(/\[!img:\$(.*?)\$\]/g, (_, key) => {
      const src = imgMap[key];
      if (!src) return "";
      return `<img class="docx-img" src="${src}" alt="${key}">`;
    });

    // math tokens -> latex, fallback to image
    html = html.replace(/\[!m:\$(.*?)\$\]/g, (_, key) => {
      const tex = latexMap[key];
      if (tex && String(tex).trim()) return `\\(${tex}\\)`;
      const fallback = imgMap["fallback_" + key];
      if (fallback) return `<img class="inline-img" src="${fallback}" alt="${key}">`;
      return "";
    });

    return html;
  }

  function tfPill(v) {
    if (v === true) return `<span class="tf-pill pill-true">Đ</span>`;
    if (v === false) return `<span class="tf-pill pill-false">S</span>`;
    return `<span class="tf-pill pill-unknown">?</span>`;
  }

  function handleFileSelect(event) {
    selectedFile = event.target.files[0];
    const label = document.getElementById('fileLabel');
    const uploadBtn = document.getElementById('uploadBtn');

    if (selectedFile) {
      label.classList.add('has-file');
      label.innerHTML = `
        <div style="font-size:3em;margin-bottom:10px;">✅</div>
        <div style="font-size:1.2em;font-weight:bold;margin-bottom:5px;">${selectedFile.name}</div>
        <div style="color:#666;">${(selectedFile.size / 1024 / 1024).toFixed(2)} MB</div>
      `;
      uploadBtn.disabled = false;
    }
  }

  async function uploadFile() {
    if (!selectedFile) return alert("Vui lòng chọn file!");

    const loading = document.getElementById('loading');
    const results = document.getElementById('results');
    const uploadSection = document.querySelector('.upload-section');

    uploadSection.style.display = 'none';
    loading.style.display = 'block';
    results.style.display = 'none';

    try {
      const fd = new FormData();
      fd.append('file', selectedFile);

      const res = await fetch(`${API_URL}/upload`, { method: 'POST', body: fd });
      if (!res.ok) throw new Error(`HTTP error ${res.status}`);

      const data = await res.json();
      if (!data.ok) throw new Error(data.error || "Unknown error");

      latexMap = data.latex || {};
      imgMap = data.images || {};
      rawText = data.rawText || "";
      questions = data.questions || [];

      // build debug
      const debug = {
        total: data.total ?? questions.length,
        mcq: questions.filter(q => q.type === "mcq" || q.type === "multiple_choice").length,
        tf4: questions.filter(q => q.type === "tf4").length,
        short: questions.filter(q => q.type === "short").length,
        latexKeys: Object.keys(latexMap).length,
        latexOk: Object.values(latexMap).filter(v => v && String(v).trim()).length,
        imageKeys: Object.keys(imgMap).length
      };

      displayStats(debug);
      displayOverview(debug);
      displayQuestions();
      displayLatexDebug(debug);
      displayHtmlPreview();
      displayDebugInfo(debug);
      displayRawJson({ ok:true, ...data });

      loading.style.display = 'none';
      results.style.display = 'block';

      setTimeout(() => window.MathJax && MathJax.typesetPromise(), 50);

    } catch (e) {
      loading.style.display = 'none';
      results.style.display = 'block';
      results.innerHTML = `
        <div class="error">
          <h3>❌ Lỗi khi xử lý file</h3>
          <p><strong>Chi tiết:</strong> ${escapeHtml(e.message)}</p>
          <button onclick="location.reload()" style="margin-top:15px;padding:10px 20px;background:#721c24;color:white;border:none;border-radius:5px;cursor:pointer;">
            🔄 Thử Lại
          </button>
        </div>
      `;
    }
  }

  function displayStats(dbg) {
    const el = document.getElementById('statsGrid');
    el.innerHTML = `
      <div class="stat-card">
        <div class="stat-value">${dbg.total}</div>
        <div class="stat-label">Tổng Câu Hỏi</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${dbg.mcq}</div>
        <div class="stat-label">Trắc Nghiệm</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${dbg.tf4}</div>
        <div class="stat-label">Đúng/Sai</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${dbg.latexOk}</div>
        <div class="stat-label">Công Thức LaTeX</div>
      </div>
    `;
  }

  function displayOverview(dbg) {
    const el = document.getElementById('tab-overview');
    el.innerHTML = `
      <div class="debug-info">
        <h3>📊 Thống Kê Chi Tiết</h3>
        <div class="debug-grid">
          <div class="debug-item"><label>Tổng câu hỏi</label><div class="value">${dbg.total}</div></div>
          <div class="debug-item"><label>MCQ</label><div class="value">${dbg.mcq}</div></div>
          <div class="debug-item"><label>TF4</label><div class="value">${dbg.tf4}</div></div>
          <div class="debug-item"><label>Short</label><div class="value">${dbg.short}</div></div>
          <div class="debug-item"><label>LaTeX OK</label><div class="value">${dbg.latexOk}</div></div>
          <div class="debug-item"><label>Ảnh keys</label><div class="value">${dbg.imageKeys}</div></div>
        </div>
      </div>
      ${dbg.total === 0 ? `
        <div class="error" style="margin-top:20px;">
          <h3>⚠️ Không parse được câu hỏi</h3>
          <p>Regex parsing server chưa match định dạng file Word của bạn.</p>
        </div>
      ` : ""}
    `;
  }

  function displayQuestions() {
    const tab = document.getElementById('tab-questions');
    if (!questions.length) {
      tab.innerHTML = `<div class="error"><h3>❌ Không có câu hỏi</h3></div>`;
      return;
    }

    let html = "";

    questions.forEach((q, idx) => {
      const no = q.no ?? (idx + 1);
      const type = q.type === "multiple_choice" ? "mcq" : (q.type || "mcq");
      const typeLabel = { mcq:"Trắc Nghiệm", tf4:"Đúng/Sai", short:"Tự Luận" }[type] || type;

      html += `
        <div class="question-card">
          <div class="question-header">
            <span class="question-number">Câu ${no}</span>
            <span class="question-type type-${type}">${typeLabel}</span>
          </div>
          <div class="question-stem">${renderInline(q.content || q.stem || "")}</div>
      `;

      if (type === "mcq") {
        const choices = q.choices || q.choicesHtml || [];
        // server mới có q.choices = [{label,text}]
        let map = {};
        if (Array.isArray(choices)) {
          for (const c of choices) map[c.label] = renderInline(c.text || "");
        } else {
          map = choices; // already map
        }
        const ans = q.correct || q.answer || null;

        html += `<div class="choices">`;
        ["A","B","C","D"].forEach(k => {
          const isCorrect = ans === k;
          html += `
            <div class="choice ${isCorrect ? "correct" : ""}">
              <span class="choice-label">${k}.</span>
              ${map[k] || ""}
              ${isCorrect ? " ✅" : ""}
            </div>
          `;
        });
        html += `</div>`;
      }

      if (type === "tf4") {
        const st = q.statements || {};
        const ansObj = q.answer || {}; // {a:true/false/null...}

        html += `<div class="choices">`;
        ["a","b","c","d"].forEach(k => {
          html += `
            <div class="choice ${ansObj[k] === true ? "correct" : ""}">
              <span class="choice-label">${k})</span>
              ${renderInline(st[k] || "")}
              ${tfPill(ansObj[k])}
            </div>
          `;
        });
        html += `</div>`;
      }

      if (type === "short") {
        html += `
          <div style="margin-top:10px">
            <div class="answer-input-wrapper">
              <label class="answer-input-label">Đáp án:</label>
              <input type="text" class="answer-input" placeholder="Nhập đáp án..." />
            </div>
          </div>
        `;
      }

      if (q.solution) {
        html += `
          <button class="solution-toggle-btn" onclick="toggleSolution(this)">👁️ Xem lời giải</button>
          <div class="solution-content" style="display:none;">
            <div class="solution-section">${renderInline(q.solution)}</div>
          </div>
        `;
      }

      html += `</div>`;
    });

    tab.innerHTML = html;
    setTimeout(() => window.MathJax && MathJax.typesetPromise([tab]), 10);
  }

  function displayLatexDebug(dbg) {
    const tab = document.getElementById('tab-latex');
    const keys = Object.keys(latexMap || {});
    const hasEmfWmf = Object.values(imgMap || {}).some(v =>
      typeof v === "string" && (v.startsWith("data:image/emf") || v.startsWith("data:image/wmf"))
    );

    let items = "";
    keys.slice(0, 30).forEach(k => {
      const tex = latexMap[k] || "";
      const fallback = imgMap["fallback_" + k] || "";
      items += `
        <div style="margin-top:12px;padding:12px;background:white;border-radius:10px;border-left:4px solid ${tex ? '#28a745':'#667eea'};">
          <div style="font-weight:bold;color:#667eea;margin-bottom:8px;">${k} ${tex ? '✅ latex' : '⚠️ fallback'}</div>
          <pre style="background:#f8f9fa;padding:10px;border-radius:8px;white-space:pre-wrap;word-break:break-word;max-height:180px;overflow:auto;">${escapeHtml(tex || "(empty)")}</pre>
          <div style="font-size:1.2em;background:#e7f3ff;padding:12px;border-radius:8px;overflow:auto;">
            ${tex ? `\\(${tex}\\)` : (fallback ? `<img class="inline-img" src="${fallback}" alt="${k}">` : '(no fallback)')}
          </div>
        </div>
      `;
    });

    tab.innerHTML = `
      <div class="debug-info">
        <h3>🔬 LaTeX Debug</h3>
        <div class="debug-grid">
          <div class="debug-item"><label>Tổng công thức</label><div class="value">${keys.length}</div></div>
          <div class="debug-item"><label>LaTeX OK</label><div class="value">${dbg.latexOk}</div></div>
          <div class="debug-item"><label>EMF/WMF còn?</label><div class="value" style="color:${hasEmfWmf? '#dc3545':'#28a745'}">${hasEmfWmf ? 'YES' : 'NO'}</div></div>
        </div>
        <p style="margin-top:12px;color:#666;">Ưu tiên hiển thị LaTeX. Nếu latex rỗng → fallback ảnh <code>fallback_mathtype_x</code>.</p>
      </div>
      ${items || '<div class="debug-info"><p>Không có công thức</p></div>'}
    `;

    setTimeout(() => window.MathJax && MathJax.typesetPromise([tab]), 10);
  }

  function displayHtmlPreview() {
    const tab = document.getElementById('tab-html');

    // build preview from questions
    let html = "";
    questions.forEach((q, i) => {
      html += `<div class="question-card">`;
      html += `<div class="question-stem"><b>Câu ${q.no ?? (i+1)}.</b> ${renderInline(q.content || q.stem || "")}</div>`;
      if (q.type === "tf4" && q.statements) {
        ["a","b","c","d"].forEach(k => {
          html += `<div class="choice"><span class="choice-label">${k})</span> ${renderInline(q.statements[k]||"")}</div>`;
        });
      } else if (Array.isArray(q.choices)) {
        q.choices.forEach(c => {
          html += `<div class="choice"><span class="choice-label">${c.label}.</span> ${renderInline(c.text||"")}</div>`;
        });
      }
      if (q.solution) html += `<div class="solution-content" style="display:block;margin-top:10px"><div class="solution-section">${renderInline(q.solution)}</div></div>`;
      html += `</div>`;
    });

    tab.innerHTML = `<div class="html-preview">${html || "<p>Không có HTML</p>"}</div>`;
    setTimeout(() => window.MathJax && MathJax.typesetPromise([tab]), 10);
  }

  function displayDebugInfo(dbg) {
    const tab = document.getElementById('tab-debug');
    tab.innerHTML = `
      <div class="debug-info">
        <h3>🔧 Debug Info</h3>
        <pre style="background:#f8f9fa;padding:20px;border-radius:10px;overflow:auto;white-space:pre-wrap;word-break:break-word;">${escapeHtml(JSON.stringify(dbg, null, 2))}</pre>
      </div>
      <div class="debug-info">
        <h3>🧾 Raw Text (server)</h3>
        <pre style="background:#0b1020;color:#d7e1ff;padding:20px;border-radius:10px;overflow:auto;max-height:280px;white-space:pre-wrap;word-break:break-word;">${escapeHtml(rawText || "")}</pre>
      </div>
    `;
  }

  function syntaxHighlight(json) {
    json = json.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    return json.replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g,
      function(match){
        let cls='number';
        if(/^"/.test(match)){ cls=/:$/.test(match)?'key':'string'; }
        else if(/true|false/.test(match)){ cls='boolean'; }
        else if(/null/.test(match)){ cls='null'; }
        const color = {key:'#e06c75',string:'#98c379',number:'#d19a66',boolean:'#56b6c2',null:'#c678dd'}[cls];
        return '<span style="color:'+color+'">'+match+'</span>';
      }
    );
  }

  function displayRawJson(data) {
    const tab = document.getElementById('tab-json');
    tab.innerHTML = `<div class="json-view">${syntaxHighlight(JSON.stringify(data, null, 2))}</div>`;
  }

  function toggleSolution(button) {
    const solutionContent = button.nextElementSibling;
    if (!solutionContent || !solutionContent.classList.contains('solution-content')) return;
    const open = solutionContent.style.display === 'none' || solutionContent.style.display === '';
    solutionContent.style.display = open ? 'block' : 'none';
    button.innerHTML = open ? '🙈 Ẩn lời giải' : '👁️ Xem lời giải';
    button.classList.toggle('active', open);
    if (open && window.MathJax) MathJax.typesetPromise([solutionContent]);
  }

  function switchTab(tabName, btnEl) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    if (btnEl) btnEl.classList.add('active');

    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    document.getElementById('tab-' + tabName).classList.add('active');

    if (window.MathJax) MathJax.typesetPromise();
  }

  // Drag & drop
  const fileLabel = document.getElementById('fileLabel');
  const fileInput = document.getElementById('fileInput');
  ['dragenter','dragover','dragleave','drop'].forEach(eventName => {
    fileLabel.addEventListener(eventName, preventDefaults, false);
  });
  function preventDefaults(e){ e.preventDefault(); e.stopPropagation(); }
  ['dragenter','dragover'].forEach(eventName => {
    fileLabel.addEventListener(eventName, () => {
      fileLabel.style.borderColor='#764ba2';
      fileLabel.style.background='#e9ecef';
    });
  });
  ['dragleave','drop'].forEach(eventName => {
    fileLabel.addEventListener(eventName, () => {
      fileLabel.style.borderColor='#667eea';
      fileLabel.style.background='#f8f9fa';
    });
  });
  fileLabel.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    fileInput.files = files;
    handleFileSelect({ target: { files } });
  });
</script>
</body>
</html>
