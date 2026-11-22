// QY Office Viewer - Frontend Logic
// Communicates with Sidecar backend & Antigravity host bridge

const sidecar = window.sidecar;
const $ = (id) => document.getElementById(id);

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let currentFile = null;
let sniffedFiles = [];
let currentZoom = 100;
let isAutoFit = true;
let currentDocNaturalWidth = 816;
let currentPdfDoc = null;
let currentPdfNaturalWidth = 595;
let currentWorkbook = null;
let activeSheetName = null;
let toastTimer = null;

// URL & Query Routing State
let lastKnownSearch = '';
let lastKnownQueryFile = null;
let lastKnownQueryTimestamp = null;

// Smart On-Demand Interaction & Dynamic Heartbeat State
let isCheckingModification = false;
let lastCheckTime = 0;
let lastActivityTime = Date.now();
let heartbeatTimer = null;

// ---------------------------------------------------------------------------
// Helpers & Toasts
// ---------------------------------------------------------------------------
function toast(message, type = 'info') {
  const el = $('toast');
  el.textContent = message;
  el.className = `toast ${type}`;
  el.style.display = 'flex';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    el.style.display = 'none';
  }, 3200);
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function formatDate(timestamp) {
  if (!timestamp) return '';
  const d = new Date(timestamp);
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

async function callApi(path, { method = 'GET', body } = {}) {
  const res = await sidecar.fetch(path, {
    method,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `${res.status} ${res.statusText}`);
  }
  return data;
}

function getUrlParams() {
  try {
    const params = new URLSearchParams(window.location.search);
    const file = params.get('file') || params.get('path') || null;
    const timestamp = params.get('_t') || params.get('t') || null;
    return { file, timestamp };
  } catch {
    return { file: null, timestamp: null };
  }
}

function getUrlFileParam() {
  return getUrlParams().file;
}

function normPath(p) {
  return (p || '').replace(/\\/g, '/').toLowerCase();
}

// ---------------------------------------------------------------------------
// Host Theme Synchronization & Preservation
// ---------------------------------------------------------------------------
function syncAndPreserveTheme() {
  try {
    const rootStyle = getComputedStyle(document.documentElement);
    const primaryVal = rootStyle.getPropertyValue('--primary').trim();
    const bgVal = rootStyle.getPropertyValue('--background').trim();

    if (primaryVal && primaryVal !== '#2563eb') {
      localStorage.setItem('qy_theme_primary', primaryVal);
      if (bgVal) localStorage.setItem('qy_theme_bg', bgVal);
    } else {
      const savedPrimary = localStorage.getItem('qy_theme_primary');
      const savedBg = localStorage.getItem('qy_theme_bg');
      if (savedPrimary) {
        document.documentElement.style.setProperty('--primary', savedPrimary);
        document.documentElement.style.setProperty('--primary-color', savedPrimary);
      }
      if (savedBg) {
        document.documentElement.style.setProperty('--background', savedBg);
        document.documentElement.style.setProperty('--bg-color', savedBg);
      }
    }
  } catch {}
}

// ---------------------------------------------------------------------------
// Document Route & Synchronizer (Always Fresh & Dynamic Auto-Watcher)
// ---------------------------------------------------------------------------
async function syncWithUrl(force = false) {
  const { file: urlFile, timestamp: urlTimestamp } = getUrlParams();
  lastKnownSearch = window.location.search || '';
  lastKnownQueryFile = urlFile;
  lastKnownQueryTimestamp = urlTimestamp;

  if (!urlFile) {
    if (!currentFile) {
      showEmptyState('未指定预览文档，请在对话中点击文档胶囊打开');
      updateButtonsState(null);
      updateMetaDisplay(null);
      $('header-status').textContent = '就绪';
    }
    return;
  }

  const isSameFile = currentFile && normPath(currentFile.path) === normPath(urlFile);

  let meta = null;
  try {
    meta = await callApi(`/api/file/meta?path=${encodeURIComponent(urlFile)}&_t=${Date.now()}`);
  } catch (err) {
    console.warn('[syncWithUrl] Meta query notice:', err);
  }

  // 同一文件若非强制刷新，检查磁盘时间戳与文件大小是否变化
  if (isSameFile && !force) {
    if (meta && currentFile.mtime && meta.mtime === currentFile.mtime && meta.size === currentFile.size) {
      return; // 磁盘无改动，保持当前视图
    }
  }

  const effectivePath = meta?.path || urlFile;
  const name = meta?.name || effectivePath.split(/[\/\\]/).pop();
  const ext = meta?.ext || ('.' + name.split('.').pop().toLowerCase());
  let fileInfo = {
    name,
    path: effectivePath,
    originalRequestedPath: urlFile,
    ext,
    size: meta?.size || 0,
    mtime: meta?.mtime || Date.now(),
  };

  notifyUserActivity(false);
  await loadFile(fileInfo);
}

// 重新加载当前正在查看的文档
async function reloadCurrentDocument(showToastNotice = true) {
  if (!currentFile || !currentFile.path) return;
  try {
    const meta = await callApi(`/api/file/meta?path=${encodeURIComponent(currentFile.path)}&_t=${Date.now()}`);
    if (meta && !meta.error) {
      currentFile.size = meta.size;
      currentFile.mtime = meta.mtime;
    }
  } catch (err) {
    console.warn('[reloadCurrentDocument] meta fetch notice:', err);
  }
  await loadFile(currentFile);
  if (showToastNotice) {
    toast(`已重新载入最新文档`, 'success');
  }
}

// ---------------------------------------------------------------------------
// Smart On-Demand Modification Checker & Adaptive Heartbeat
// ---------------------------------------------------------------------------
function notifyUserActivity(doCheckImmediately = true) {
  lastActivityTime = Date.now();
  if (doCheckImmediately) {
    checkAndReloadIfModified();
  }
}

async function checkAndReloadIfModified(force = false) {
  if (!currentFile || !currentFile.path) return;
  const now = Date.now();
  if (!force && now - lastCheckTime < 600) return; // 节流 600ms
  if (isCheckingModification) return;

  lastCheckTime = now;
  isCheckingModification = true;

  try {
    const meta = await callApi(`/api/file/meta?path=${encodeURIComponent(currentFile.path)}&_t=${now}`);
    if (meta && !meta.error) {
      const isMtimeChanged = currentFile.mtime && meta.mtime && meta.mtime !== currentFile.mtime;
      const isSizeChanged = currentFile.size !== undefined && meta.size !== undefined && meta.size !== currentFile.size;

      if (isMtimeChanged || isSizeChanged) {
        console.log('[Auto-Reload] 物理文件已更新，自动热重载最新文档:', meta);
        currentFile.size = meta.size;
        currentFile.mtime = meta.mtime;
        await loadFile(currentFile);
        toast('文档已更新，已自动载入最新内容', 'success');
      }
    }
  } catch (err) {
    // 忽略静默探测失败
  } finally {
    isCheckingModification = false;
  }
}

// 自适应动态心跳：
// 1. 文档刚打开或用户刚操作后的前 25 秒（修改高发期）：每 2.5 秒轻量探测一次
// 2. 25-90 秒（阅读平稳期）：降频至约每 8 秒探测一次
// 3. 超过 90 秒无操作（静置期）：彻底休眠（0 次网络/IO 请求），完全依靠用户鼠标划入或胶囊点击唤醒
function startAdaptiveHeartbeat() {
  if (heartbeatTimer) clearInterval(heartbeatTimer);
  heartbeatTimer = setInterval(() => {
    if (!currentFile || !currentFile.path) return;
    const idleSeconds = (Date.now() - lastActivityTime) / 1000;
    if (idleSeconds <= 25) {
      checkAndReloadIfModified();
    } else if (idleSeconds <= 90) {
      if (Math.floor(idleSeconds) % 8 < 3) {
        checkAndReloadIfModified();
      }
    }
  }, 2500);
}


// ---------------------------------------------------------------------------
// Load and Render File
// ---------------------------------------------------------------------------
async function loadFile(fileInfo) {
  if (!fileInfo || !fileInfo.path) return;
  currentFile = fileInfo;
  updateButtonsState(fileInfo);
  updateMetaDisplay(fileInfo);

  const ext = (fileInfo.ext || '').toLowerCase();
  let typeLabel = 'Office 预览';
  if (ext === '.docx' || ext === '.doc') typeLabel = 'Word 文档';
  else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') typeLabel = 'Excel 表格';
  else if (ext === '.pdf') typeLabel = 'PDF 文档';
  $('header-status').textContent = typeLabel;

  showLoading(`正在载入 ${fileInfo.name}...`);
  try {
    if (ext === '.docx') {
      await renderDocx(fileInfo.path);
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      await renderSpreadsheet(fileInfo.path, ext);
    } else if (ext === '.pdf') {
      await renderPdf(fileInfo.path);
    } else {
      showEmptyState(`暂不支持在侧边栏直接预览 ${ext} 文件，请使用顶部按钮在本地默认程序中打开。`);
    }
  } catch (err) {
    console.error('[render] Error:', err);
    showEmptyState(`渲染出错: ${err.message}`);
    toast(`渲染失败: ${err.message}`, 'error');
  }
}

function updateButtonsState(file) {
  const hasFile = !!file;
  $('btn-open-default').disabled = !hasFile;
  $('btn-reveal').disabled = !hasFile;
  $('btn-copy-path').disabled = !hasFile;

  if (file) {
    const ext = (file.ext || '').toLowerCase();
    let label = '打开';
    if (ext === '.docx' || ext === '.doc') label = 'Word 打开';
    else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') label = 'Excel 打开';
    else if (ext === '.pdf') label = 'PDF 打开';
    $('btn-open-label').textContent = label;
    $('btn-open-default').title = `在系统默认程序中打开：${file.name}`;
    $('btn-reveal').title = `在资源管理器中定位：${file.name}`;
    $('btn-copy-path').title = `复制绝对路径：${file.path}`;
  } else {
    $('btn-open-label').textContent = '打开';
    $('btn-open-default').title = '在系统默认程序中打开';
  }
}

function updateMetaDisplay(file) {
  if (!file) {
    $('file-meta-text').textContent = '未选择文件';
    return;
  }
  const parts = [
    file.name,
    formatBytes(file.size),
    formatDate(file.mtime),
  ];
  $('file-meta-text').textContent = parts.filter(Boolean).join(' · ');
}

// ---------------------------------------------------------------------------
// Render Engines
// ---------------------------------------------------------------------------

// 1. DOCX Renderer (docx-preview with proper inWrapper & Horizontal Full-Width Fit)
async function renderDocx(filePath) {
  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
  if (!res.ok) {
    let errMsg = '无法读取文件内容 (文件可能不存在或无法访问)';
    try {
      const errJson = await res.json();
      if (errJson && errJson.error) errMsg = errJson.error;
    } catch {}
    throw new Error(errMsg);
  }
  const arrayBuffer = await res.arrayBuffer();

  hideAllViewports();
  $('docx-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'flex';

  const container = $('docx-render-target');
  container.innerHTML = '';

  if (window.docx && window.docx.renderAsync) {
    // 关键：ignoreHeight: true 保证页面高度随正文段落自然伸展，杜绝半截白底和后续文字掉落
    await window.docx.renderAsync(arrayBuffer, container, null, {
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
      experimental: true,
    });

    // 智能修正国标公文落款缩进与中心对齐（GB/T 9704-2012）
    fixDocxSignatures(container);

    // 测量未缩放状态下的自然物理宽度
    const section = container.querySelector('section.docx');
    if (section) {
      currentDocNaturalWidth = section.offsetWidth || 816;
    } else {
      currentDocNaturalWidth = 816;
    }

    // 默认开启横向排满，多层渲染帧兜底保证视口宽度测量 100% 准确
    isAutoFit = true;
    requestAnimationFrame(() => {
      fixDocxSignatures(container);
      fitDocxToWidth();
      setTimeout(() => {
        fixDocxSignatures(container);
        if (isAutoFit) fitDocxToWidth();
      }, 100);
    });
  } else {
    throw new Error('docx-preview 引擎未就绪');
  }
}

// ---------------------------------------------------------------------------
// Standard Government Signature Post-Processor (GB/T 9704-2012 Compliance)
// ---------------------------------------------------------------------------
function fixDocxSignatures(container) {
  if (!container) return;
  const paragraphs = Array.from(container.querySelectorAll('section.docx p, .docx-wrapper p, p'));
  if (paragraphs.length === 0) return;

  // 正则：成文日期格式（支持阿拉伯数字、大写汉字、零/〇等全角半角）
  const dateRegex = /^[〇一二三四五六七八九十零0-9]{4}\s*年\s*[一二三四五六七八九十0-9]{1,2}\s*月\s*[一二三四五六七八九十0-9]{1,2}\s*日$/;
  // 机构/落款关键词常见结尾
  const orgSuffixRegex = /(?:公司|党委|办公室|纪委|工会|委员会|支部|部门|部|处|科|局|组|团委|支公司|分公司|中心)$/;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (p.classList.contains('gov-signature-line')) continue;

    const rawText = (p.innerText || p.textContent || '').trim();
    const cleanText = rawText.replace(/[\s\t\u3000\u00a0]+/g, '');

    if (dateRegex.test(cleanText)) {
      // 找到了成文日期段落！向上回溯收集落款组（发文机关/部门署名）
      const sigGroup = [p];
      let j = i - 1;
      while (j >= 0 && sigGroup.length < 5) {
        const prevP = paragraphs[j];
        if (prevP.classList.contains('gov-signature-line')) break;
        const prevText = (prevP.innerText || prevP.textContent || '').trim();
        const prevClean = prevText.replace(/[\s\t\u3000\u00a0]+/g, '');

        if (!prevClean) {
          j--;
          continue;
        }

        // 遇到附件或标题或长正文，停止回溯
        if (prevClean.startsWith('附件') || prevClean.includes('附件：') || prevClean.includes('附件:')) {
          break;
        }

        const hasTab = prevText.includes('\t') || !!prevP.querySelector('.docx-tab-stop');
        const isOrgLike = prevClean.length <= 35 && (orgSuffixRegex.test(prevClean) || hasTab);

        if (isOrgLike || (hasTab && prevClean.length <= 40)) {
          sigGroup.unshift(prevP);
          j--;
        } else {
          break;
        }
      }

      // 计算这一组落款中最长一行的视觉字符宽度（汉字算1，非汉字算0.55）
      let maxVisualWidth = 0;
      const groupData = sigGroup.map((sigP) => {
        const text = (sigP.innerText || sigP.textContent || '').replace(/[\t\r\n]+/g, '').trim();
        let w = 0;
        for (const char of text) {
          w += char.charCodeAt(0) > 255 ? 1 : 0.55;
        }
        if (w > maxVisualWidth) maxVisualWidth = w;
        return { p: sigP, text, visualWidth: w };
      });

      if (maxVisualWidth >= 2) {
        const targetWidthEm = Math.ceil(maxVisualWidth) + 0.5;
        groupData.forEach(({ p: sigP, text }) => {
          sigP.classList.add('gov-signature-line');
          sigP.style.setProperty('text-align', 'right', 'important');
          sigP.style.setProperty('padding-right', '4em', 'important'); // 国标GB/T 9704-2012：右空四字
          sigP.style.setProperty('margin-left', '0', 'important');
          sigP.style.setProperty('text-indent', '0', 'important');
          sigP.style.setProperty('white-space', 'nowrap', 'important');

          // 构建居中包裹容器
          const block = document.createElement('span');
          block.className = 'gov-signature-block';
          block.style.setProperty('display', 'inline-block', 'important');
          block.style.setProperty('min-width', `${targetWidthEm}em`, 'important');
          block.style.setProperty('text-align', 'center', 'important');
          block.style.setProperty('white-space', 'nowrap', 'important');

          // 收集内部子节点，去除多余制表符与前导空白
          const validNodes = [];
          Array.from(sigP.childNodes).forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList.contains('docx-tab-stop') || node.innerHTML === '&emsp;' || node.textContent === '\t') {
                return;
              }
            } else if (node.nodeType === Node.TEXT_NODE) {
              if (node.textContent.trim() === '') return;
              node.textContent = node.textContent.replace(/^[\t\s]+/, '');
            }
            validNodes.push(node);
          });

          sigP.innerHTML = '';
          if (validNodes.length > 0) {
            validNodes.forEach((n) => block.appendChild(n));
          } else {
            block.textContent = text;
          }
          sigP.appendChild(block);
        });
      }
    }
  }

  // 补充扫描：仅有发文机关而无日期的单独落款行（以较多空格/制表符开头且以常见机构结尾）
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (p.classList.contains('gov-signature-line')) continue;
    const raw = (p.innerText || p.textContent || '');
    if ((/^[\t\s\u3000\u00a0]{4,}/.test(raw) || p.querySelector('.docx-tab-stop')) && raw.trim().length <= 30) {
      const clean = raw.trim().replace(/[\s\t\u3000\u00a0]+/g, '');
      if (orgSuffixRegex.test(clean)) {
        p.classList.add('gov-signature-line');
        p.style.setProperty('text-align', 'right', 'important');
        p.style.setProperty('padding-right', '4em', 'important');
        p.style.setProperty('margin-left', '0', 'important');
        p.style.setProperty('text-indent', '0', 'important');
        p.style.setProperty('white-space', 'nowrap', 'important');

        const block = document.createElement('span');
        block.className = 'gov-signature-block';
        block.style.setProperty('display', 'inline-block', 'important');
        block.style.setProperty('text-align', 'center', 'important');
        block.textContent = raw.trim();
        p.innerHTML = '';
        p.appendChild(block);
      }
    }
  }
}

// Fit docx to viewport width (严格按横向排满设置比例)
function fitDocxToWidth() {
  const container = $('docx-render-target');
  const section = container.querySelector('section.docx');
  const viewport = $('viewport-container');
  if (!section || !viewport) return;

  const innerWrapper = container.querySelector('.docx-wrapper') || container;

  // 获得文档未缩放时的自然宽度
  if (!currentDocNaturalWidth || currentDocNaturalWidth <= 0) {
    const prevZoom = innerWrapper.style.zoom;
    innerWrapper.style.zoom = '1';
    currentDocNaturalWidth = section.offsetWidth || 816;
    if (prevZoom) innerWrapper.style.zoom = prevZoom;
  }

  // viewport.clientWidth 扣除纵向滚动条后的纯可用宽度
  const availableWidth = viewport.clientWidth;

  if (docWidthAvailable(availableWidth, currentDocNaturalWidth)) {
    const scale = availableWidth / currentDocNaturalWidth;
    applyDocxZoom(scale);
    isAutoFit = true;
  }
}

function docWidthAvailable(w, nat) {
  return w > 0 && nat > 0;
}

function applyDocxZoom(scale) {
  currentZoom = Math.round(scale * 100);
  $('zoom-label').textContent = `${currentZoom}%`;

  const container = $('docx-render-target');
  const innerWrapper = container.querySelector('.docx-wrapper') || container;

  if ('zoom' in innerWrapper.style) {
    innerWrapper.style.zoom = scale;
  } else {
    innerWrapper.style.transform = `scale(${scale})`;
    innerWrapper.style.transformOrigin = 'top center';
  }
}

// 2. Spreadsheet Renderer (Backend Lossless CJK HTML Stream + Fallback)
let currentSheetsMap = {};

async function renderSpreadsheet(filePath, ext) {
  hideAllViewports();
  showLoading(`正在载入表格 ${currentFile ? currentFile.name : ''}...`);

  try {
    // 优先调用后端高保真无损解析 API（由 Node.js 服务端完成 CJK/CP936 解码，100% 杜绝浏览器 latin1 乱码）
    const res = await callApi(`/api/file/spreadsheet?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
    if (!res || !res.ok) {
      throw new Error((res && res.error) || '服务端解析表格失败');
    }

    hideAllViewports();
    $('xlsx-wrapper').style.display = 'flex';
    $('zoom-controls').style.display = 'none';

    const sheetNames = res.sheetNames || [];
    if (sheetNames.length === 0) {
      $('xlsx-table-wrapper').innerHTML = '<div style="padding:20px;text-align:center;">工作簿为空</div>';
      return;
    }

    currentSheetsMap = res.sheets || {};

    const tabsContainer = $('sheet-tabs');
    tabsContainer.innerHTML = '';
    sheetNames.forEach((name, idx) => {
      const tab = document.createElement('div');
      tab.className = `sheet-tab ${idx === 0 ? 'active' : ''}`;
      tab.textContent = name;
      tab.addEventListener('click', () => {
        document.querySelectorAll('.sheet-tab').forEach((t) => t.classList.remove('active'));
        tab.classList.add('active');
        renderSheetTable(name);
      });
      tabsContainer.appendChild(tab);
    });

    renderSheetTable(sheetNames[0]);
  } catch (backendErr) {
    console.warn('[renderSpreadsheet] Backend parse failed, trying client fallback:', backendErr);
    // 降级使用前端本地 SheetJS
    try {
      const rawRes = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
      if (!rawRes.ok) throw new Error('无法读取表格原始数据');
      const arrayBuffer = await rawRes.arrayBuffer();

      hideAllViewports();
      $('xlsx-wrapper').style.display = 'flex';
      $('zoom-controls').style.display = 'none';

      if (!window.XLSX) throw new Error('XLSX 引擎未就绪');
      const data = new Uint8Array(arrayBuffer);
      currentWorkbook = window.XLSX.read(data, { type: 'array' });
      const sheetNames = currentWorkbook.SheetNames || [];
      if (sheetNames.length === 0) {
        $('xlsx-table-wrapper').innerHTML = '<div style="padding:20px;text-align:center;">工作簿为空</div>';
        return;
      }

      currentSheetsMap = {};
      for (const sname of sheetNames) {
        currentSheetsMap[sname] = window.XLSX.utils.sheet_to_html(currentWorkbook.Sheets[sname], {
          id: 'sheet-table',
          editable: false,
        });
      }

      const tabsContainer = $('sheet-tabs');
      tabsContainer.innerHTML = '';
      sheetNames.forEach((name, idx) => {
        const tab = document.createElement('div');
        tab.className = `sheet-tab ${idx === 0 ? 'active' : ''}`;
        tab.textContent = name;
        tab.addEventListener('click', () => {
          document.querySelectorAll('.sheet-tab').forEach((t) => t.classList.remove('active'));
          tab.classList.add('active');
          renderSheetTable(name);
        });
        tabsContainer.appendChild(tab);
      });

      renderSheetTable(sheetNames[0]);
    } catch (clientErr) {
      console.error('[renderSpreadsheet] Fallback failed:', clientErr);
      showEmptyState(`表格解析失败: ${backendErr.message}`);
      toast(`表格解析失败: ${backendErr.message}`, 'error');
    }
  }
}

function renderSheetTable(sheetName) {
  activeSheetName = sheetName;
  const html = currentSheetsMap[sheetName] || '';
  const tableContainer = $('xlsx-table-wrapper');
  tableContainer.innerHTML = html;
  const table = tableContainer.querySelector('table');
  if (table) {
    table.className = 'sheet-table';
    table.querySelectorAll('td').forEach((td) => {
      if (td.textContent.includes('\n') || td.textContent.length > 40) {
        td.classList.add('multiline-cell');
      }
    });
  }
}

// 3. PDF Renderer (Mozilla PDF.js Offline High-DPI Multi-Page Canvas Stream)
async function ensurePdfEngine() {
  if (window.pdfjsLib) return;
  for (let i = 0; i < 6; i++) {
    if (window.pdfjsLib) return;
    await new Promise((r) => setTimeout(r, 60));
  }
  if (window.pdfjsLib) return;

  try {
    const r1 = await sidecar.fetch('/vendor/pdf.min.js').catch(() => null);
    if (r1 && r1.ok) {
      const scriptText = await r1.text();
      const script = document.createElement('script');
      script.textContent = scriptText;
      document.head.appendChild(script);
      await new Promise((r) => setTimeout(r, 80));
    }
  } catch (err) {
    console.error('[PDF.js] Dynamic engine load error:', err);
  }

  if (!window.pdfjsLib) {
    throw new Error('PDF 离线引擎 (pdf.min.js) 尚未完成加载');
  }
}

async function ensurePdfWorker() {
  if (!window.pdfjsLib) return;
  if (pdfjsLib.GlobalWorkerOptions.workerPort || pdfjsLib.GlobalWorkerOptions.workerSrc) return;

  try {
    const testR = await sidecar.fetch('/vendor/pdf.worker.min.js').catch(() => null);
    if (testR && testR.ok) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = '/vendor/pdf.worker.min.js';
      return;
    }
  } catch {}
}

async function renderPdf(filePath) {
  hideAllViewports();
  showLoading(`正在载入 PDF ${currentFile ? currentFile.name : ''}...`);

  await ensurePdfEngine();
  await ensurePdfWorker();

  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
  if (!res.ok) {
    let errMsg = '无法读取 PDF 文件数据 (文件可能不存在或无法访问)';
    try {
      const errJson = await res.json();
      if (errJson && errJson.error) errMsg = errJson.error;
    } catch {}
    throw new Error(errMsg);
  }
  const arrayBuffer = await res.arrayBuffer();

  const loadingTask = pdfjsLib.getDocument({
    data: arrayBuffer,
    cMapPacked: true,
  });

  const pdfDoc = await loadingTask.promise;
  currentPdfDoc = pdfDoc;
  const numPages = pdfDoc.numPages;

  hideAllViewports();
  $('pdf-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'flex';

  const container = $('pdf-pages-container');
  container.innerHTML = '';

  const firstPage = await pdfDoc.getPage(1);
  const baseViewport = firstPage.getViewport({ scale: 1.0 });
  currentPdfNaturalWidth = baseViewport.width || 595;

  const dpr = window.devicePixelRatio || 1;
  const renderScale = Math.max(1.5, Math.min(2.5, dpr * 1.25));

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = pageNum === 1 ? firstPage : await pdfDoc.getPage(pageNum);
    const pageViewport1x = page.getViewport({ scale: 1.0 });
    const renderViewport = page.getViewport({ scale: renderScale });

    const pageCard = document.createElement('div');
    pageCard.className = 'pdf-page-card';
    pageCard.setAttribute('data-page-number', pageNum);

    const canvas = document.createElement('canvas');
    canvas.className = 'pdf-page-canvas';
    canvas.width = Math.floor(renderViewport.width);
    canvas.height = Math.floor(renderViewport.height);
    canvas.style.width = `${Math.floor(pageViewport1x.width)}px`;
    canvas.style.height = `${Math.floor(pageViewport1x.height)}px`;
    pageCard.appendChild(canvas);

    const pageBadge = document.createElement('div');
    pageBadge.className = 'pdf-page-badge';
    pageBadge.textContent = `${pageNum} / ${numPages}`;
    pageCard.appendChild(pageBadge);

    container.appendChild(pageCard);

    const ctx = canvas.getContext('2d');
    await page.render({
      canvasContext: ctx,
      viewport: renderViewport,
    }).promise;
  }

  isAutoFit = true;
  requestAnimationFrame(() => {
    fitPdfToWidth();
    setTimeout(() => {
      if (isAutoFit) fitPdfToWidth();
    }, 120);
  });
}

function fitPdfToWidth() {
  const viewport = $('viewport-container');
  const container = $('pdf-pages-container');
  if (!viewport || !container || !currentPdfNaturalWidth) return;

  const availableWidth = viewport.clientWidth - 20;
  if (availableWidth > 0 && currentPdfNaturalWidth > 0) {
    const scale = availableWidth / currentPdfNaturalWidth;
    applyPdfZoom(scale);
    isAutoFit = true;
  }
}

function applyPdfZoom(scale) {
  currentZoom = Math.round(scale * 100);
  $('zoom-label').textContent = `${currentZoom}%`;

  const container = $('pdf-pages-container');
  if (!container) return;

  if ('zoom' in container.style) {
    container.style.zoom = scale;
  } else {
    container.style.transform = `scale(${scale})`;
    container.style.transformOrigin = 'top center';
  }
}

// ---------------------------------------------------------------------------
// Viewport State Switches
// ---------------------------------------------------------------------------
function hideAllViewports() {
  $('empty-state').style.display = 'none';
  $('loading-state').style.display = 'none';
  $('docx-wrapper').style.display = 'none';
  $('xlsx-wrapper').style.display = 'none';
  const pdfWrapper = $('pdf-wrapper');
  if (pdfWrapper) pdfWrapper.style.display = 'none';
  const pdfIframe = $('pdf-iframe');
  if (pdfIframe) pdfIframe.style.display = 'none';
  $('text-view').style.display = 'none';
}

function showEmptyState(msg) {
  hideAllViewports();
  $('empty-state').style.display = 'flex';
  $('zoom-controls').style.display = 'none';
  if (msg) {
    const p = $('empty-state').querySelector('p:first-child');
    if (p) p.textContent = msg;
  }
}

function showLoading(msg) {
  hideAllViewports();
  $('loading-state').style.display = 'flex';
  $('loading-text').textContent = msg || '正在加载文档渲染...';
}

// ---------------------------------------------------------------------------
// Zoom Controls (DOCX & PDF)
// ---------------------------------------------------------------------------
function updateZoom(newPercent) {
  isAutoFit = false;
  const scale = Math.max(0.2, Math.min(3.0, newPercent / 100));
  if ($('docx-wrapper').style.display !== 'none') {
    applyDocxZoom(scale);
  } else if ($('pdf-wrapper') && $('pdf-wrapper').style.display !== 'none') {
    applyPdfZoom(scale);
  }
}

function resetZoom() {
  isAutoFit = false;
  if ($('docx-wrapper').style.display !== 'none') {
    applyDocxZoom(1.0);
  } else if ($('pdf-wrapper') && $('pdf-wrapper').style.display !== 'none') {
    applyPdfZoom(1.0);
  }
}

function fitCurrentToWidth() {
  isAutoFit = true;
  if ($('docx-wrapper').style.display !== 'none') {
    fitDocxToWidth();
  } else if ($('pdf-wrapper') && $('pdf-wrapper').style.display !== 'none') {
    fitPdfToWidth();
  }
}

// ---------------------------------------------------------------------------
// Native System Actions (P0 - Instant Sub-10ms Silent Dispatch)
// ---------------------------------------------------------------------------
async function openInDefaultApp() {
  if (!currentFile || !currentFile.path) return;
  try {
    const res = await callApi('/api/action/open-default', {
      method: 'POST',
      body: { path: currentFile.path },
    });
    if (!res.ok) {
      toast(res.error || '打开失败', 'error');
    }
  } catch (err) {
    toast(`启动失败: ${err.message}`, 'error');
  }
}

async function revealInExplorer() {
  if (!currentFile || !currentFile.path) return;
  try {
    const res = await callApi('/api/action/reveal', {
      method: 'POST',
      body: { path: currentFile.path },
    });
    if (!res.ok) {
      toast(res.error || '定位失败', 'error');
    }
  } catch (err) {
    toast(`定位失败: ${err.message}`, 'error');
  }
}

// Robust clipboard copy with fallback to native OS command via sidecar
async function copyToClipboard(text) {
  // 1. Try traditional execCommand (bypasses Permissions Policy in iframes)
  let copied = false;
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.style.top = '0';
    ta.setAttribute('readonly', '');
    document.body.appendChild(ta);
    ta.select();
    ta.setSelectionRange(0, 99999);
    copied = document.execCommand('copy');
    document.body.removeChild(ta);
  } catch {
    copied = false;
  }

  if (copied) return true;

  // 2. Fallback to Sidecar backend native OS clipboard API
  try {
    const res = await callApi('/api/action/copy-clipboard', {
      method: 'POST',
      body: { text },
    });
    return !!res.ok;
  } catch (err) {
    console.error('[copy-clipboard] Error:', err);
    return false;
  }
}

async function copyFilePath() {
  if (!currentFile || !currentFile.path) return;
  const ok = await copyToClipboard(currentFile.path);
  if (ok) {
    toast('文件绝对路径已复制到剪贴板', 'success');
  } else {
    toast('复制失败，请在上方手动选择路径复制', 'error');
  }
}

// ---------------------------------------------------------------------------
// Drag and Drop & Manual Path
// ---------------------------------------------------------------------------
function setupDragAndDrop() {
  const overlay = $('drop-overlay');

  window.addEventListener('dragover', (e) => {
    e.preventDefault();
    overlay.style.display = 'flex';
  });

  window.addEventListener('dragleave', (e) => {
    if (e.target === overlay || e.clientX === 0 || e.clientY === 0) {
      overlay.style.display = 'none';
    }
  });

  window.addEventListener('drop', async (e) => {
    e.preventDefault();
    overlay.style.display = 'none';

    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      const file = files[0];
      const name = file.name;
      const localPath = file.path || '';
      if (localPath) {
        const ext = '.' + name.split('.').pop().toLowerCase();
        loadFile({
          name,
          path: localPath,
          size: file.size,
          mtime: file.lastModified,
          ext,
          source: 'manual',
        });
        toast(`已载入文件：${name}`, 'info');
      } else {
        renderRawDroppedFile(file);
      }
    }
  });
}

async function renderRawDroppedFile(file) {
  const ext = '.' + file.name.split('.').pop().toLowerCase();
  showLoading(`正在读取本地文件 ${file.name}...`);
  currentFile = {
    name: file.name,
    path: file.name,
    size: file.size,
    mtime: file.lastModified,
    ext,
  };
  updateButtonsState(null);
  updateMetaDisplay(currentFile);

  try {
    const arrayBuffer = await file.arrayBuffer();
    if (ext === '.docx') {
      hideAllViewports();
      $('docx-wrapper').style.display = 'flex';
      $('zoom-controls').style.display = 'flex';
      $('docx-render-target').innerHTML = '';
      await window.docx.renderAsync(arrayBuffer, $('docx-render-target'), null, {
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: true,
        renderHeaders: true,
        renderFooters: true,
      });
      const sec = $('docx-render-target').querySelector('section.docx');
      currentDocNaturalWidth = sec ? (sec.offsetWidth || 816) : 816;
      isAutoFit = true;
      requestAnimationFrame(() => fitDocxToWidth());
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      hideAllViewports();
      $('xlsx-wrapper').style.display = 'flex';
      $('zoom-controls').style.display = 'none';
      const data = new Uint8Array(arrayBuffer);
      currentWorkbook = window.XLSX.read(data, { type: 'array' });
      renderSheetTable(currentWorkbook.SheetNames[0]);
    } else if (ext === '.pdf') {
      if (!window.pdfjsLib) throw new Error('PDF 离线引擎未就绪');
      if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = '/vendor/pdf.worker.min.js';
      }
      const pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer, cMapPacked: true }).promise;
      currentPdfDoc = pdfDoc;
      hideAllViewports();
      $('pdf-wrapper').style.display = 'flex';
      $('zoom-controls').style.display = 'flex';
      const container = $('pdf-pages-container');
      container.innerHTML = '';
      const firstPage = await pdfDoc.getPage(1);
      const baseViewport = firstPage.getViewport({ scale: 1.0 });
      currentPdfNaturalWidth = baseViewport.width || 595;
      const dpr = window.devicePixelRatio || 1;
      const renderScale = Math.max(1.5, Math.min(2.5, dpr * 1.25));

      for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
        const page = pageNum === 1 ? firstPage : await pdfDoc.getPage(pageNum);
        const pageViewport1x = page.getViewport({ scale: 1.0 });
        const renderViewport = page.getViewport({ scale: renderScale });
        const pageCard = document.createElement('div');
        pageCard.className = 'pdf-page-card';
        pageCard.setAttribute('data-page-number', pageNum);
        const canvas = document.createElement('canvas');
        canvas.className = 'pdf-page-canvas';
        canvas.width = Math.floor(renderViewport.width);
        canvas.height = Math.floor(renderViewport.height);
        canvas.style.width = `${Math.floor(pageViewport1x.width)}px`;
        canvas.style.height = `${Math.floor(pageViewport1x.height)}px`;
        pageCard.appendChild(canvas);
        const pageBadge = document.createElement('div');
        pageBadge.className = 'pdf-page-badge';
        pageBadge.textContent = `${pageNum} / ${pdfDoc.numPages}`;
        pageCard.appendChild(pageBadge);
        container.appendChild(pageCard);
        await page.render({ canvasContext: canvas.getContext('2d'), viewport: renderViewport }).promise;
      }
      isAutoFit = true;
      requestAnimationFrame(() => fitPdfToWidth());
    } else {
      showEmptyState(`暂不支持直接脱机解析 ${ext}`);
    }
  } catch (err) {
    showEmptyState(`解析失败: ${err.message}`);
  }
}

// ---------------------------------------------------------------------------
// Event Listeners Initialization
// ---------------------------------------------------------------------------
function initEvents() {
  $('btn-open-default').addEventListener('click', openInDefaultApp);
  $('btn-reveal').addEventListener('click', revealInExplorer);
  $('btn-copy-path').addEventListener('click', copyFilePath);

  // Document Reload button (In-place instant silent refresh, never drops host CSS variables)
  const reloadUiBtn = $('btn-reload-ui');
  if (reloadUiBtn) {
    reloadUiBtn.addEventListener('click', async () => {
      syncAndPreserveTheme();
      await reloadCurrentDocument(true);
    });
  }

  // Zoom controls
  $('btn-zoom-in').addEventListener('click', () => updateZoom(currentZoom + 10));
  $('btn-zoom-out').addEventListener('click', () => updateZoom(currentZoom - 10));
  $('btn-zoom-reset').addEventListener('click', resetZoom);
  $('btn-zoom-fit').addEventListener('click', () => {
    fitCurrentToWidth();
  });

  setupDragAndDrop();

  // Watch for dynamic resize of the auxiliary side pane
  if (window.ResizeObserver) {
    const ro = new ResizeObserver(() => {
      if (isAutoFit && currentFile) {
        if (currentFile.ext === '.docx') {
          fitDocxToWidth();
        } else if (currentFile.ext === '.pdf') {
          fitPdfToWidth();
        }
      }
    });
    ro.observe($('viewport-container'));
  }

  // URL query parameter watcher & Window focus trigger (Pill clicks refocus the panel)
  window.addEventListener('popstate', () => syncWithUrl(true));
  window.addEventListener('hashchange', () => syncWithUrl(true));

  // 接收 Antigravity 宿主通过 postMessage 派发的消息
  window.addEventListener('message', () => {
    notifyUserActivity(true);
  });

  // 每次窗口重新获得焦点（例如用户在会话中点击文档胶囊切换过来），强制同步并检查最新文件
  window.addEventListener('focus', () => {
    syncAndPreserveTheme();
    notifyUserActivity(true);
  });

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) {
      syncAndPreserveTheme();
      notifyUserActivity(true);
    }
  });

  // 鼠标划入或交互感知（用户在左侧点完胶囊，鼠标/目光移向右侧侧边栏即可零等待感知最新文件）
  ['pointerenter', 'mouseenter', 'pointerdown', 'wheel'].forEach((evt) => {
    window.addEventListener(evt, () => notifyUserActivity(true), { passive: true });
  });

  // 高频微秒级 URL search 变动监听（纯本地内存字符串比对，零网络请求、零磁盘IO）
  setInterval(() => {
    const currentSearch = window.location.search || '';
    if (currentSearch !== lastKnownSearch) {
      lastKnownSearch = currentSearch;
      syncWithUrl(true);
    }
  }, 250);

  // 启动自适应平滑心跳（活跃期 2.5s 探测，静置 90s 后彻底休眠）
  startAdaptiveHeartbeat();
}

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  syncAndPreserveTheme();
  initEvents();
  setTimeout(() => {
    syncAndPreserveTheme();
    lastKnownSearch = window.location.search || '';
    lastKnownQueryFile = getUrlFileParam();
    syncWithUrl(true);
  }, 50);
});
