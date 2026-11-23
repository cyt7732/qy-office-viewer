// QY Office Viewer - Frontend Logic
// Communicates with Sidecar backend & Antigravity host bridge

const sidecar = window.sidecar;
const $ = (id) => document.getElementById(id);
const CLIENT_PLATFORM = navigator.userAgentData?.platform || navigator.platform || navigator.userAgent || '';
const IS_WINDOWS = /win/i.test(CLIENT_PLATFORM);
const IS_MAC = /mac/i.test(CLIENT_PLATFORM);

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let currentFile = null;
let currentZoom = 100;
let isAutoFit = true;
let currentDocNaturalWidth = 816;
let currentPdfNaturalWidth = 595;
let activeSheetName = null;
let currentSheetsMap = {};
let toastTimer = null;

// Concurrency cancellation token (prevents interleaved/stale renders)
let activeRenderToken = 0;

// URL & Query Routing State
let lastKnownSearch = '';

// Adaptive Watch State Check
let isCheckingModification = false;
let lastCheckTime = 0;
let lastActivityTime = Date.now();
let heartbeatTimer = null;

// Script Lazy-Loader Cache
const loadedScripts = new Map();

// PDF Lazy & TextLayer State
const pdfState = {
  doc: null,
  observer: null,
  pages: [],
  scale: 1.0,
  zoomDebounceTimer: null,
  visiblePages: new Set(),
};

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
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
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

function normPath(p) {
  const normalized = (p || '').replace(/\\/g, '/');
  return IS_WINDOWS ? normalized.toLowerCase() : normalized;
}

function getBaseName(p) {
  return (p || '').split(/[\/\\]/).pop() || '';
}

function clampScale(scale) {
  return Math.max(0.25, Math.min(3.0, scale));
}

function decodeTextBuffer(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  if (bytes.length === 0) return '';

  // UTF-8 BOM
  if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }
  // UTF-16LE BOM
  if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  }
  // UTF-16BE BOM
  if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
    const payload = bytes.subarray(2);
    const evenLen = payload.length - (payload.length % 2);
    const swapped = new Uint8Array(evenLen);
    for (let i = 0; i < evenLen; i += 2) {
      swapped[i] = payload[i + 1];
      swapped[i + 1] = payload[i];
    }
    return new TextDecoder('utf-16le').decode(swapped);
  }

  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('gb18030').decode(bytes);
  }
}

// ---------------------------------------------------------------------------
// On-Demand Vendor Script Loader
// ---------------------------------------------------------------------------
function loadScriptOnce(src) {
  if (loadedScripts.has(src)) {
    return loadedScripts.get(src);
  }

  const promise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => {
      loadedScripts.delete(src);
      reject(new Error(`加载离线渲染引擎失败: ${src}`));
    };
    document.head.appendChild(script);
  });

  loadedScripts.set(src, promise);
  return promise;
}

async function ensureDocxEngine() {
  if (!window.JSZip) {
    await loadScriptOnce('/vendor/jszip.min.js');
  }
  if (!window.docx || !window.docx.renderAsync) {
    await loadScriptOnce('/vendor/docx-preview.min.js');
  }
  if (!window.docx || !window.docx.renderAsync) {
    throw new Error('docx-preview 引擎未就绪');
  }
}

async function ensureXlsxEngine() {
  if (!window.XLSX) {
    await loadScriptOnce('/vendor/xlsx.full.min.js');
  }
  if (!window.XLSX) {
    throw new Error('XLSX 引擎未就绪');
  }
}

async function ensurePdfEngine() {
  if (!window.pdfjsLib) {
    await loadScriptOnce('/vendor/pdf.min.js');
  }
  if (!window.pdfjsLib) {
    throw new Error('PDF 离线引擎 (pdf.min.js) 尚未完成加载');
  }
  if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/vendor/pdf.worker.min.js';
  }
}

// ---------------------------------------------------------------------------
// View State Capture & Restore (Hot-Reload Scroll/Sheet Preservation)
// ---------------------------------------------------------------------------
function getActiveScrollElement() {
  if ($('docx-wrapper').style.display !== 'none') return $('viewport-container');
  if ($('pdf-wrapper').style.display !== 'none') return $('pdf-wrapper');
  if ($('xlsx-wrapper').style.display !== 'none') return $('xlsx-table-wrapper');
  if ($('text-view').style.display !== 'none') return $('text-view');
  return null;
}

function captureViewState() {
  const scrollEl = getActiveScrollElement();
  const scrollTop = scrollEl ? scrollEl.scrollTop : 0;
  const scrollLeft = scrollEl ? scrollEl.scrollLeft : 0;
  const maxScroll = scrollEl ? Math.max(0, scrollEl.scrollHeight - scrollEl.clientHeight) : 0;
  const scrollRatio = maxScroll > 0 ? scrollTop / maxScroll : 0;

  return {
    filePath: currentFile ? normPath(currentFile.path) : null,
    scrollTop,
    scrollLeft,
    scrollRatio,
    activeSheetName,
    isAutoFit,
    scale: currentZoom / 100,
  };
}

function restoreScrollPosition(scrollEl, viewState) {
  if (!scrollEl || !viewState) return;
  const applyScroll = () => {
    const maxScroll = Math.max(0, scrollEl.scrollHeight - scrollEl.clientHeight);
    if (maxScroll > 0 && viewState.scrollRatio > 0) {
      scrollEl.scrollTop = Math.min(maxScroll, Math.round(viewState.scrollRatio * maxScroll));
    } else if (viewState.scrollTop > 0) {
      scrollEl.scrollTop = Math.min(maxScroll, viewState.scrollTop);
    }
    if (viewState.scrollLeft > 0) {
      scrollEl.scrollLeft = viewState.scrollLeft;
    }
  };
  applyScroll();
  requestAnimationFrame(applyScroll);
}

// ---------------------------------------------------------------------------
// Document Route & Synchronizer
// ---------------------------------------------------------------------------
async function syncWithUrl(force = false) {
  const { file: urlFile } = getUrlParams();
  lastKnownSearch = window.location.search || '';

  if (!urlFile) {
    if (!currentFile) {
      showEmptyState(
        '未指定预览文档',
        '请在对话中点击文档胶囊打开，或将本地文件拖入此处'
      );
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

  if (meta && meta.exists === false) {
    showMissingFileState(urlFile, meta.candidates || []);
    return;
  }

  if (meta && meta.allowed === false) {
    cleanupPdfState();
    currentFile = null;
    updateButtonsState(null);
    updateMetaDisplay(null);
    $('header-status').textContent = '不支持';
    showEmptyState(`暂不支持预览 ${meta.ext || '该'} 类型文件`, '仅支持常见 Office 与文本格式预览');
    return;
  }

  if (isSameFile && !force) {
    if (meta && currentFile.mtime && meta.mtime === currentFile.mtime && meta.size === currentFile.size) {
      return;
    }
  }

  const effectivePath = meta?.path || urlFile;
  const name = meta?.name || getBaseName(effectivePath);
  const ext = meta?.ext || (name.includes('.') ? '.' + name.split('.').pop().toLowerCase() : '');

  let watchVersion = currentFile?.watchVersion ?? 0;
  try {
    const ws = await callApi(`/api/file/watch-state?path=${encodeURIComponent(effectivePath)}&_t=${Date.now()}`);
    if (ws && typeof ws.version === 'number') watchVersion = ws.version;
  } catch {
    // ignore initial watch-state error
  }

  const fileInfo = {
    name,
    path: effectivePath,
    originalRequestedPath: urlFile,
    ext,
    size: meta?.size || 0,
    mtime: meta?.mtime || Date.now(),
    watchVersion,
  };

  const preserveViewState = isSameFile ? captureViewState() : null;
  markUserActivity();
  await loadFile(fileInfo, { preserveViewState });
}

async function reloadCurrentDocument(showToastNotice = true) {
  if (!currentFile || !currentFile.path || currentFile.isMemoryOnly) return;

  const preserveViewState = captureViewState();

  try {
    const meta = await callApi(`/api/file/meta?path=${encodeURIComponent(currentFile.path)}&_t=${Date.now()}`);
    if (meta && meta.exists === false) {
      showMissingFileState(currentFile.path, meta.candidates || []);
      return;
    }
    if (meta && meta.exists) {
      currentFile.size = meta.size;
      currentFile.mtime = meta.mtime;
    }
    const ws = await callApi(`/api/file/watch-state?path=${encodeURIComponent(currentFile.path)}&_t=${Date.now()}`);
    if (ws && typeof ws.version === 'number') {
      currentFile.watchVersion = ws.version;
    }
  } catch (err) {
    console.warn('[reloadCurrentDocument] meta fetch notice:', err);
  }

  await loadFile(currentFile, { preserveViewState });
  if (showToastNotice) {
    toast('已重新载入最新文档', 'success');
  }
}

// ---------------------------------------------------------------------------
// Adaptive File Change Watcher (Powered by Backend fs.watch + watch-state)
// ---------------------------------------------------------------------------
function markUserActivity() {
  lastActivityTime = Date.now();
}

async function checkAndReloadIfModified(force = false) {
  if (!currentFile || !currentFile.path || currentFile.isMemoryOnly) return;
  if (!force && document.hidden) return;

  const now = Date.now();
  if (!force && now - lastCheckTime < 1000) return;
  if (isCheckingModification) return;

  lastCheckTime = now;
  isCheckingModification = true;

  try {
    const state = await callApi(`/api/file/watch-state?path=${encodeURIComponent(currentFile.path)}&_t=${now}`);
    if (!state) return;

    if (state.exists === false) {
      const missingPath = currentFile.path;
      let candidates = [];
      try {
        const meta = await callApi(`/api/file/meta?path=${encodeURIComponent(missingPath)}&_t=${now}`);
        candidates = meta?.candidates || [];
      } catch {}
      showMissingFileState(missingPath, candidates);
      return;
    }

    const versionChanged = typeof state.version === 'number' && state.version !== currentFile.watchVersion;
    const mtimeChanged = state.mtime && state.mtime !== currentFile.mtime;
    const sizeChanged = state.size !== undefined && state.size !== currentFile.size;

    if (versionChanged || mtimeChanged || sizeChanged) {
      const preserveViewState = captureViewState();
      currentFile.size = state.size;
      currentFile.mtime = state.mtime;
      currentFile.watchVersion = state.version ?? currentFile.watchVersion;
      await loadFile(currentFile, { preserveViewState });
      toast('文档已更新，已自动载入最新内容', 'success');
    }
  } catch {
    // ignore silent probe failure
  } finally {
    isCheckingModification = false;
  }
}

function startAdaptiveHeartbeat() {
  if (heartbeatTimer) clearInterval(heartbeatTimer);
  let tick = 0;
  heartbeatTimer = setInterval(() => {
    if (document.hidden || !currentFile || !currentFile.path || currentFile.isMemoryOnly) return;
    tick++;
    const idleSeconds = (Date.now() - lastActivityTime) / 1000;
    if (idleSeconds <= 30) {
      checkAndReloadIfModified();
    } else if (tick % 3 === 0) {
      // Every ~9-10s when reading calmly while visible
      checkAndReloadIfModified();
    }
  }, 3000);
}

// ---------------------------------------------------------------------------
// Load and Render File
// ---------------------------------------------------------------------------
async function loadFile(fileInfo, { preserveViewState = null } = {}) {
  if (!fileInfo || !fileInfo.path) return;
  const token = ++activeRenderToken;

  cleanupPdfState();

  currentFile = fileInfo;
  updateButtonsState(fileInfo);
  updateMetaDisplay(fileInfo);

  const ext = (fileInfo.ext || '').toLowerCase();
  let typeLabel = 'Office 预览';
  if (ext === '.docx' || ext === '.doc') typeLabel = 'Word 文档';
  else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') typeLabel = 'Excel 表格';
  else if (ext === '.pdf') typeLabel = 'PDF 文档';
  else if (ext === '.txt' || ext === '.md' || ext === '.json') typeLabel = '文本预览';
  $('header-status').textContent = typeLabel;

  showLoading(`正在载入 ${fileInfo.name}...`);

  try {
    if (ext === '.docx') {
      await renderDocx(fileInfo.path, token, preserveViewState);
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      await renderSpreadsheet(fileInfo.path, ext, token, preserveViewState);
    } else if (ext === '.pdf') {
      await renderPdf(fileInfo.path, token, preserveViewState);
    } else if (ext === '.txt' || ext === '.md' || ext === '.json') {
      await renderTextDocument(fileInfo.path, ext, token, preserveViewState);
    } else {
      showEmptyState(
        `暂不支持在侧边栏直接预览 ${ext} 文件`,
        '请使用顶部「打开」按钮在本地默认程序中打开。'
      );
    }
  } catch (err) {
    if (token !== activeRenderToken) return;
    console.error('[render] Error:', err);
    showEmptyState(`渲染出错: ${err.message}`);
    toast(`渲染失败: ${err.message}`, 'error');
  }
}

function updateButtonsState(file) {
  const hasRealPath = !!(file && file.path && !file.isMemoryOnly);
  $('btn-open-default').disabled = !hasRealPath;
  $('btn-reveal').disabled = !hasRealPath;
  $('btn-copy-path').disabled = !hasRealPath;

  if (file) {
    const ext = (file.ext || '').toLowerCase();
    let label = '打开';
    if (ext === '.docx' || ext === '.doc') label = 'Word 打开';
    else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') label = 'Excel 打开';
    else if (ext === '.pdf') label = 'PDF 打开';
    $('btn-open-label').textContent = label;

    if (hasRealPath) {
      const revealLabel = IS_WINDOWS ? '资源管理器' : (IS_MAC ? '访达（Finder）' : '文件管理器');
      $('btn-open-default').title = `在系统默认程序中打开：${file.name}`;
      $('btn-reveal').title = `在${revealLabel}中定位：${file.name}`;
      $('btn-copy-path').title = `复制绝对路径：${file.path}`;
    } else {
      $('btn-open-default').title = '拖放的内存文件无本地完整路径';
      $('btn-reveal').title = '拖放的内存文件无本地完整路径';
      $('btn-copy-path').title = '拖放的内存文件无本地完整路径';
    }
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
// 1. DOCX Renderer
// ---------------------------------------------------------------------------
async function renderDocx(filePath, token, preserveViewState = null) {
  await ensureDocxEngine();
  if (token !== activeRenderToken) return;

  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
  if (token !== activeRenderToken) return;

  if (!res.ok) {
    let errMsg = '无法读取文件内容 (文件可能不存在或无法访问)';
    try {
      const errJson = await res.json();
      if (errJson && errJson.error) errMsg = errJson.error;
    } catch {}
    throw new Error(errMsg);
  }
  const arrayBuffer = await res.arrayBuffer();
  if (token !== activeRenderToken) return;

  await mountDocxBuffer(arrayBuffer, token, preserveViewState);
}

async function mountDocxBuffer(arrayBuffer, token, preserveViewState = null) {
  await ensureDocxEngine();
  if (token !== activeRenderToken) return;

  hideAllViewports();
  $('docx-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'flex';

  const container = $('docx-render-target');
  container.innerHTML = '';

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

  if (token !== activeRenderToken) return;

  // Single-pass conservative signature alignment (never touches tables/headers/footers/body)
  fixDocxSignatures(container);

  const section = container.querySelector('section.docx');
  currentDocNaturalWidth = section ? (section.offsetWidth || 816) : 816;

  if (preserveViewState && !preserveViewState.isAutoFit && preserveViewState.scale > 0) {
    isAutoFit = false;
    applyDocxZoom(clampScale(preserveViewState.scale));
  } else {
    isAutoFit = true;
    requestAnimationFrame(() => {
      if (token !== activeRenderToken) return;
      fitDocxToWidth();
    });
  }

  if (preserveViewState) {
    restoreScrollPosition($('viewport-container'), preserveViewState);
  } else {
    $('viewport-container').scrollTop = 0;
  }
}

// Conservative Government Signature Post-Processor (GB/T 9704-2012)
function fixDocxSignatures(container) {
  if (!container) return;
  const bodyRoot = container.querySelector('section.docx > article') || container.querySelector('section.docx') || container;

  // Exclude paragraphs inside tables, headers, footers, or lists
  const allEligible = Array.from(bodyRoot.querySelectorAll('p')).filter(
    (p) => !p.closest('table, header, footer, li')
  );
  if (allEligible.length === 0) return;

  // Only inspect the trailing tail of the document (last 15 non-empty paragraphs)
  const nonEmptyTail = allEligible
    .filter((p) => (p.textContent || '').replace(/[\s\u3000\u00a0]+/g, '').length > 0)
    .slice(-15);

  const dateRegex = /^[〇一二三四五六七八九十零0-9]{4}\s*年\s*[一二三四五六七八九十0-9]{1,2}\s*月\s*[一二三四五六七八九十0-9]{1,2}\s*日$/;
  const sentencePunctRegex = /[，。、；：？！,.;:?!]/;
  const orgSuffixRegex = /(?:公司|党委|办公室|纪委|工会|委员会|支部|部门|部|处|科|局|组|团委|支公司|分公司|中心)$/;

  for (let i = nonEmptyTail.length - 1; i >= 0; i--) {
    const p = nonEmptyTail[i];
    if (p.classList.contains('gov-signature-line')) continue;

    const rawText = (p.innerText || p.textContent || '').trim();
    if (sentencePunctRegex.test(rawText)) continue;

    const cleanText = rawText.replace(/[\s\t\u3000\u00a0]+/g, '');
    if (!dateRegex.test(cleanText)) continue;

    const sigGroup = [p];
    let j = i - 1;
    while (j >= 0 && sigGroup.length < 4) {
      const prevP = nonEmptyTail[j];
      if (prevP.classList.contains('gov-signature-line')) break;

      const prevRaw = (prevP.innerText || prevP.textContent || '').trim();
      const prevClean = prevRaw.replace(/[\s\t\u3000\u00a0]+/g, '');

      if (
        !prevClean ||
        prevClean.length > 35 ||
        /[。？！；：?!;:]/.test(prevRaw) ||
        prevClean.startsWith('附件') ||
        prevClean.includes('附件：') ||
        prevClean.includes('附件:')
      ) {
        break;
      }

      const hasTab = prevRaw.includes('\t') || !!prevP.querySelector('.docx-tab-stop');
      if (orgSuffixRegex.test(prevClean) || (hasTab && prevClean.length <= 35)) {
        sigGroup.unshift(prevP);
        j--;
      } else {
        break;
      }
    }

    let maxVisualWidth = 0;
    const groupData = sigGroup.map((sigP) => {
      const text = (sigP.innerText || sigP.textContent || '').replace(/[\t\r\n]+/g, '').trim();
      let w = 0;
      for (const char of text) {
        w += char.charCodeAt(0) > 255 ? 1 : 0.55;
      }
      if (w > maxVisualWidth) maxVisualWidth = w;
      return { p: sigP, text };
    });

    if (maxVisualWidth >= 2) {
      const targetWidthEm = Math.ceil(maxVisualWidth) + 0.5;
      groupData.forEach(({ p: sigP, text }) => {
        sigP.classList.add('gov-signature-line');
        sigP.style.setProperty('text-align', 'right', 'important');
        sigP.style.setProperty('padding-right', '4em', 'important');
        sigP.style.setProperty('margin-left', '0', 'important');
        sigP.style.setProperty('text-indent', '0', 'important');

        const block = document.createElement('span');
        block.className = 'gov-signature-block';
        block.style.setProperty('display', 'inline-block', 'important');
        block.style.setProperty('min-width', `${targetWidthEm}em`, 'important');
        block.style.setProperty('text-align', 'center', 'important');

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

    // Process only the final signature block at the end of the document
    break;
  }
}

function fitDocxToWidth() {
  const container = $('docx-render-target');
  const section = container.querySelector('section.docx');
  const viewport = $('viewport-container');
  if (!section || !viewport) return;

  const innerWrapper = container.querySelector('.docx-wrapper') || container;

  if (!currentDocNaturalWidth || currentDocNaturalWidth <= 0) {
    const prevZoom = innerWrapper.style.zoom;
    innerWrapper.style.zoom = '1';
    currentDocNaturalWidth = section.offsetWidth || 816;
    if (prevZoom) innerWrapper.style.zoom = prevZoom;
  }

  const availableWidth = viewport.clientWidth;
  if (availableWidth > 0 && currentDocNaturalWidth > 0) {
    applyDocxZoom(clampScale(availableWidth / currentDocNaturalWidth));
    isAutoFit = true;
  }
}

function applyDocxZoom(scale) {
  const safeScale = clampScale(scale);
  currentZoom = Math.round(safeScale * 100);
  $('zoom-label').textContent = `${currentZoom}%`;

  const container = $('docx-render-target');
  const innerWrapper = container.querySelector('.docx-wrapper') || container;

  if ('zoom' in innerWrapper.style) {
    innerWrapper.style.zoom = safeScale;
  } else {
    innerWrapper.style.transform = `scale(${safeScale})`;
    innerWrapper.style.transformOrigin = 'top center';
  }
}

// ---------------------------------------------------------------------------
// 2. Spreadsheet Renderer (Backend Lossless + Client Fallback + Sticky Header)
// ---------------------------------------------------------------------------
async function renderSpreadsheet(filePath, ext, token, preserveViewState = null) {
  try {
    const res = await callApi(`/api/file/spreadsheet?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
    if (token !== activeRenderToken) return;

    if (!res || !res.ok) {
      throw new Error((res && res.error) || '服务端解析表格失败');
    }

    mountSpreadsheetView(
      res.sheetNames || [],
      res.sheets || {},
      preserveViewState?.activeSheetName || null,
      preserveViewState
    );
  } catch (backendErr) {
    if (token !== activeRenderToken) return;
    console.warn('[renderSpreadsheet] Backend parse failed, trying client fallback:', backendErr);

    await ensureXlsxEngine();
    if (token !== activeRenderToken) return;

    const rawRes = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
    if (token !== activeRenderToken) return;
    if (!rawRes.ok) throw new Error('无法读取表格原始数据');

    const arrayBuffer = await rawRes.arrayBuffer();
    if (token !== activeRenderToken) return;

    const wb = ext === '.csv'
      ? window.XLSX.read(decodeTextBuffer(arrayBuffer), { type: 'string' })
      : window.XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

    const { sheetNames, sheetsMap } = extractSheetsMapFromWorkbook(wb);
    mountSpreadsheetView(sheetNames, sheetsMap, preserveViewState?.activeSheetName || null, preserveViewState);
  }
}

function extractSheetsMapFromWorkbook(wb) {
  const sheetNames = wb.SheetNames || [];
  const sheetsMap = {};
  for (const sname of sheetNames) {
    const ws = wb.Sheets[sname];
    if (ws) {
      sheetsMap[sname] = window.XLSX.utils.sheet_to_html(ws, {
        header: '',
        footer: '',
        id: 'sheet-table',
        editable: false,
      });
    }
  }
  return { sheetNames, sheetsMap };
}

function mountSpreadsheetView(sheetNames, sheetsMap, preferredSheetName = null, preserveViewState = null) {
  hideAllViewports();
  $('xlsx-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'none';

  const tabsContainer = $('sheet-tabs');
  tabsContainer.innerHTML = '';

  if (!sheetNames || sheetNames.length === 0) {
    currentSheetsMap = {};
    activeSheetName = null;
    $('xlsx-table-wrapper').innerHTML = '<div style="padding:20px;text-align:center;">工作簿为空</div>';
    return;
  }

  currentSheetsMap = sheetsMap || {};
  const targetSheet = (preferredSheetName && sheetNames.includes(preferredSheetName))
    ? preferredSheetName
    : sheetNames[0];

  sheetNames.forEach((name) => {
    const tab = document.createElement('div');
    tab.className = `sheet-tab ${name === targetSheet ? 'active' : ''}`;
    tab.textContent = name;
    tab.addEventListener('click', () => {
      tabsContainer.querySelectorAll('.sheet-tab').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');
      renderSheetTable(name);
    });
    tabsContainer.appendChild(tab);
  });

  renderSheetTable(targetSheet);

  if (preserveViewState) {
    restoreScrollPosition($('xlsx-table-wrapper'), preserveViewState);
  }
}

function promoteFirstRowToThead(table) {
  if (!table) return;
  if (table.querySelector('thead th')) return;

  const firstRow = table.querySelector('tbody tr') || table.querySelector('tr');
  if (!firstRow || firstRow.children.length === 0) return;

  const thead = document.createElement('thead');
  const headerTr = document.createElement('tr');
  for (const attr of firstRow.attributes) {
    headerTr.setAttribute(attr.name, attr.value);
  }

  Array.from(firstRow.children).forEach((cell) => {
    const th = document.createElement('th');
    for (const attr of cell.attributes) {
      th.setAttribute(attr.name, attr.value);
    }
    while (cell.firstChild) {
      th.appendChild(cell.firstChild);
    }
    headerTr.appendChild(th);
  });

  thead.appendChild(headerTr);
  firstRow.remove();
  table.insertBefore(thead, table.firstChild);
}

function renderSheetTable(sheetName) {
  activeSheetName = sheetName;
  const html = currentSheetsMap[sheetName] || '';
  const tableContainer = $('xlsx-table-wrapper');
  tableContainer.innerHTML = html;
  const table = tableContainer.querySelector('table');
  if (table) {
    table.className = 'sheet-table';
    promoteFirstRowToThead(table);
    table.querySelectorAll('td').forEach((td) => {
      if (td.textContent.includes('\n') || td.textContent.length > 40) {
        td.classList.add('multiline-cell');
      }
    });
  }
  tableContainer.scrollTop = 0;
  tableContainer.scrollLeft = 0;
}

// ---------------------------------------------------------------------------
// 3. PDF Renderer (Offline CMap + IntersectionObserver Lazy + TextLayer + Sharp Zoom)
// ---------------------------------------------------------------------------
class SidecarCMapReaderFactory {
  constructor({ baseUrl = '/api/vendor/cmap?file=', isCompressed = true } = {}) {
    this.baseUrl = baseUrl || '/api/vendor/cmap?file=';
    this.isCompressed = isCompressed;
  }

  async fetch({ name }) {
    if (!name) throw new Error('CMap name must be specified.');
    const fileName = `${name}${this.isCompressed ? '.bcmap' : ''}`;
    const res = await sidecar.fetch(`${this.baseUrl}${encodeURIComponent(fileName)}`);
    if (!res.ok) {
      throw new Error(`Unable to load CMap: ${fileName} (${res.status})`);
    }
    const buf = await res.arrayBuffer();
    return {
      cMapData: new Uint8Array(buf),
      compressionType: this.isCompressed
        ? (window.pdfjsLib?.CMapCompressionType?.BINARY ?? 1)
        : (window.pdfjsLib?.CMapCompressionType?.NONE ?? 0),
    };
  }
}

function cleanupPdfState() {
  if (pdfState.zoomDebounceTimer) {
    clearTimeout(pdfState.zoomDebounceTimer);
    pdfState.zoomDebounceTimer = null;
  }
  if (pdfState.observer) {
    pdfState.observer.disconnect();
    pdfState.observer = null;
  }
  for (const meta of pdfState.pages) {
    try { meta.renderTask?.cancel?.(); } catch {}
    try { meta.textTask?.cancel?.(); } catch {}
  }
  pdfState.pages = [];
  pdfState.visiblePages.clear();

  if (pdfState.doc) {
    const oldDoc = pdfState.doc;
    pdfState.doc = null;
    oldDoc.destroy().catch(() => {});
  }

  const container = $('pdf-pages-container');
  if (container) container.innerHTML = '';
}

async function renderPdf(filePath, token, preserveViewState = null) {
  await ensurePdfEngine();
  if (token !== activeRenderToken) return;

  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
  if (token !== activeRenderToken) return;

  if (!res.ok) {
    let errMsg = '无法读取 PDF 文件数据 (文件可能不存在或无法访问)';
    try {
      const errJson = await res.json();
      if (errJson && errJson.error) errMsg = errJson.error;
    } catch {}
    throw new Error(errMsg);
  }
  const arrayBuffer = await res.arrayBuffer();
  if (token !== activeRenderToken) return;

  await mountPdfBuffer(arrayBuffer, token, preserveViewState);
}

async function mountPdfBuffer(arrayBuffer, token, preserveViewState = null) {
  await ensurePdfEngine();
  if (token !== activeRenderToken) return;

  const loadingTask = window.pdfjsLib.getDocument({
    data: arrayBuffer,
    cMapUrl: '/api/vendor/cmap?file=',
    cMapPacked: true,
    CMapReaderFactory: SidecarCMapReaderFactory,
  });

  const pdfDoc = await loadingTask.promise;
  if (token !== activeRenderToken) {
    pdfDoc.destroy().catch(() => {});
    return;
  }

  pdfState.doc = pdfDoc;
  const numPages = pdfDoc.numPages;

  hideAllViewports();
  $('pdf-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'flex';

  const container = $('pdf-pages-container');
  container.innerHTML = '';
  container.style.zoom = '';
  container.style.transform = '';

  const firstPage = await pdfDoc.getPage(1);
  if (token !== activeRenderToken) return;

  const firstViewport = firstPage.getViewport({ scale: 1.0 });
  currentPdfNaturalWidth = firstViewport.width || 595;

  const pdfWrapper = $('pdf-wrapper');
  const availableWidth = (pdfWrapper.clientWidth || $('viewport-container').clientWidth) - 24;

  let initialScale = 1.0;
  if (preserveViewState && !preserveViewState.isAutoFit && preserveViewState.scale > 0) {
    isAutoFit = false;
    initialScale = clampScale(preserveViewState.scale);
  } else {
    isAutoFit = true;
    initialScale = availableWidth > 0 && currentPdfNaturalWidth > 0
      ? clampScale(availableWidth / currentPdfNaturalWidth)
      : 1.0;
  }

  pdfState.scale = initialScale;
  currentZoom = Math.round(initialScale * 100);
  $('zoom-label').textContent = `${currentZoom}%`;

  const pages = [];
  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = pageNum === 1 ? firstPage : await pdfDoc.getPage(pageNum);
    if (token !== activeRenderToken) return;

    const vp1 = page.getViewport({ scale: 1.0 });
    const card = document.createElement('div');
    card.className = 'pdf-page-card';
    card.setAttribute('data-page-number', String(pageNum));
    card.style.width = `${Math.max(1, Math.floor(vp1.width * initialScale))}px`;
    card.style.height = `${Math.max(1, Math.floor(vp1.height * initialScale))}px`;

    const canvas = document.createElement('canvas');
    canvas.className = 'pdf-page-canvas';

    const textLayer = document.createElement('div');
    textLayer.className = 'textLayer';

    const badge = document.createElement('div');
    badge.className = 'pdf-page-badge';
    badge.textContent = `${pageNum} / ${numPages}`;

    card.appendChild(canvas);
    card.appendChild(textLayer);
    card.appendChild(badge);
    container.appendChild(card);

    pages.push({
      pageNum,
      card,
      canvas,
      textLayer,
      baseWidth: vp1.width,
      baseHeight: vp1.height,
      renderTask: null,
      textTask: null,
      renderedScale: null,
      renderingScale: null,
    });
  }

  pdfState.pages = pages;

  if (preserveViewState) {
    restoreScrollPosition(pdfWrapper, preserveViewState);
  } else {
    pdfWrapper.scrollTop = 0;
  }

  // Lazy-render visible & near-viewport pages via IntersectionObserver
  const observer = new IntersectionObserver(
    (entries) => {
      if (token !== activeRenderToken || !pdfState.doc) return;
      for (const entry of entries) {
        const pageNum = Number(entry.target.getAttribute('data-page-number'));
        if (!pageNum) continue;
        if (entry.isIntersecting) {
          pdfState.visiblePages.add(pageNum);
          const meta = pdfState.pages[pageNum - 1];
          if (meta) {
            renderPdfPage(meta, token);
          }
        } else {
          pdfState.visiblePages.delete(pageNum);
        }
      }
    },
    {
      root: pdfWrapper,
      rootMargin: '450px 0px',
      threshold: 0.01,
    }
  );

  pdfState.observer = observer;
  for (const meta of pages) {
    observer.observe(meta.card);
  }
}

async function renderPdfPage(pageMeta, token) {
  if (!pageMeta || token !== activeRenderToken || !pdfState.doc) return;
  const targetScale = pdfState.scale;

  if (pageMeta.renderedScale === targetScale || pageMeta.renderingScale === targetScale) {
    return;
  }

  try { pageMeta.renderTask?.cancel?.(); } catch {}
  try { pageMeta.textTask?.cancel?.(); } catch {}

  pageMeta.renderingScale = targetScale;

  try {
    const page = await pdfState.doc.getPage(pageMeta.pageNum);
    if (token !== activeRenderToken || !pdfState.doc || pdfState.scale !== targetScale) return;

    const displayViewport = page.getViewport({ scale: targetScale });
    const dpr = window.devicePixelRatio || 1;
    const backingRatio = Math.min(3.0, Math.max(1.25, dpr));
    const renderViewport = page.getViewport({ scale: targetScale * backingRatio });

    const { canvas, textLayer } = pageMeta;
    canvas.width = Math.max(1, Math.floor(renderViewport.width));
    canvas.height = Math.max(1, Math.floor(renderViewport.height));

    const ctx = canvas.getContext('2d');
    const renderTask = page.render({
      canvasContext: ctx,
      viewport: renderViewport,
    });
    pageMeta.renderTask = renderTask;
    await renderTask.promise;

    if (token !== activeRenderToken || !pdfState.doc || pdfState.scale !== targetScale) return;

    textLayer.innerHTML = '';
    textLayer.style.setProperty('--scale-factor', String(displayViewport.scale));
    const textTask = window.pdfjsLib.renderTextLayer({
      textContentSource: page.streamTextContent(),
      container: textLayer,
      viewport: displayViewport,
      textDivs: [],
    });
    pageMeta.textTask = textTask;
    await textTask.promise;

    if (token === activeRenderToken && pdfState.scale === targetScale) {
      pageMeta.renderedScale = targetScale;
    }
  } catch (err) {
    if (err?.name !== 'RenderingCancelledException') {
      console.warn(`[PDF] Page ${pageMeta.pageNum} render notice:`, err);
    }
  } finally {
    if (pageMeta.renderingScale === targetScale) {
      pageMeta.renderingScale = null;
    }
  }
}

function fitPdfToWidth() {
  const pdfWrapper = $('pdf-wrapper');
  const viewport = $('viewport-container');
  const widthSource = (pdfWrapper && pdfWrapper.clientWidth > 0) ? pdfWrapper.clientWidth : viewport?.clientWidth;
  if (!widthSource || !currentPdfNaturalWidth) return;

  const availableWidth = widthSource - 24;
  if (availableWidth > 0 && currentPdfNaturalWidth > 0) {
    applyPdfZoom(clampScale(availableWidth / currentPdfNaturalWidth));
    isAutoFit = true;
  }
}

function applyPdfZoom(scale) {
  const safeScale = clampScale(scale);
  if (Math.abs(safeScale - pdfState.scale) < 0.001 && pdfState.pages.length > 0) {
    currentZoom = Math.round(safeScale * 100);
    $('zoom-label').textContent = `${currentZoom}%`;
    return;
  }

  pdfState.scale = safeScale;
  currentZoom = Math.round(safeScale * 100);
  $('zoom-label').textContent = `${currentZoom}%`;

  // Immediately resize page cards so layout & existing bitmaps scale without layout jump
  for (const meta of pdfState.pages) {
    meta.card.style.width = `${Math.max(1, Math.floor(meta.baseWidth * safeScale))}px`;
    meta.card.style.height = `${Math.max(1, Math.floor(meta.baseHeight * safeScale))}px`;
    if (meta.renderedScale !== safeScale) {
      meta.renderedScale = null;
    }
  }

  // Debounce crisp high-DPI re-render of visible + adjacent pages
  clearTimeout(pdfState.zoomDebounceTimer);
  const zoomToken = activeRenderToken;
  pdfState.zoomDebounceTimer = setTimeout(() => {
    if (zoomToken !== activeRenderToken || !pdfState.doc) return;
    const targets = new Set(pdfState.visiblePages);
    for (const p of pdfState.visiblePages) {
      if (p > 1) targets.add(p - 1);
      if (p < pdfState.pages.length) targets.add(p + 1);
    }
    if (targets.size === 0 && pdfState.pages.length > 0) {
      targets.add(1);
    }
    for (const pageNum of targets) {
      const meta = pdfState.pages[pageNum - 1];
      if (meta) renderPdfPage(meta, zoomToken);
    }
  }, 160);
}

// ---------------------------------------------------------------------------
// 4. Plain Text / Markdown / JSON Renderer
// ---------------------------------------------------------------------------
async function renderTextDocument(filePath, ext, token, preserveViewState = null) {
  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}&_t=${Date.now()}`);
  if (token !== activeRenderToken) return;

  if (!res.ok) {
    let errMsg = '无法读取文件内容';
    try {
      const errJson = await res.json();
      if (errJson && errJson.error) errMsg = errJson.error;
    } catch {}
    throw new Error(errMsg);
  }

  const arrayBuffer = await res.arrayBuffer();
  if (token !== activeRenderToken) return;

  mountTextContent(decodeTextBuffer(arrayBuffer), ext, preserveViewState);
}

function mountTextContent(rawText, ext, preserveViewState = null) {
  let content = rawText;
  if (ext === '.json') {
    try {
      content = JSON.stringify(JSON.parse(rawText), null, 2);
    } catch {
      // Keep original text when invalid JSON
    }
  }

  hideAllViewports();
  $('zoom-controls').style.display = 'none';
  const tv = $('text-view');
  tv.style.display = 'block';
  tv.textContent = content;

  if (preserveViewState) {
    restoreScrollPosition(tv, preserveViewState);
  } else {
    tv.scrollTop = 0;
  }
}

// ---------------------------------------------------------------------------
// Viewport State Switches & Missing File UI
// ---------------------------------------------------------------------------
function hideAllViewports() {
  $('empty-state').style.display = 'none';
  const candEl = $('empty-candidates');
  if (candEl) {
    candEl.style.display = 'none';
    candEl.innerHTML = '';
  }
  $('loading-state').style.display = 'none';
  $('docx-wrapper').style.display = 'none';
  $('xlsx-wrapper').style.display = 'none';
  $('pdf-wrapper').style.display = 'none';
  $('text-view').style.display = 'none';
}

function showEmptyState(title, subtitle = '') {
  hideAllViewports();
  $('empty-state').style.display = 'flex';
  $('zoom-controls').style.display = 'none';
  const titleEl = $('empty-title');
  const subtitleEl = $('empty-subtitle');
  if (titleEl) titleEl.textContent = title || '未指定预览文档';
  if (subtitleEl) subtitleEl.textContent = subtitle || '';
}

function showMissingFileState(requestedPath, candidates = []) {
  activeRenderToken++;
  cleanupPdfState();
  currentFile = null;
  updateButtonsState(null);
  updateMetaDisplay(null);
  $('header-status').textContent = '文件不存在';

  const fileName = getBaseName(requestedPath) || requestedPath;
  showEmptyState(
    `文件不存在：${fileName}`,
    candidates.length > 0
      ? '原始路径未找到精确匹配文件，请在下方确认是否为重命名/相邻文件：'
      : '请确认文件路径是否已被移动、重命名或删除。'
  );

  const candEl = $('empty-candidates');
  if (!candEl || !Array.isArray(candidates) || candidates.length === 0) return;

  candEl.innerHTML = '';
  const header = document.createElement('div');
  header.className = 'missing-candidates-title';
  header.textContent = '同目录相似文件（点击打开）：';
  candEl.appendChild(header);

  candidates.forEach((cand) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'candidate-btn';

    const nameSpan = document.createElement('span');
    nameSpan.className = 'candidate-name';
    nameSpan.textContent = cand.name;
    nameSpan.title = cand.path;

    const metaSpan = document.createElement('span');
    metaSpan.className = 'candidate-meta';
    const pct = Math.round((cand.similarity || 0) * 100);
    metaSpan.textContent = `${pct}% · ${formatBytes(cand.size)}`;

    btn.appendChild(nameSpan);
    btn.appendChild(metaSpan);
    btn.addEventListener('click', () => {
      markUserActivity();
      loadFile({
        name: cand.name,
        path: cand.path,
        ext: cand.ext,
        size: cand.size,
        mtime: cand.mtime,
      });
    });

    candEl.appendChild(btn);
  });

  candEl.style.display = 'flex';
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
  const scale = clampScale(newPercent / 100);
  if ($('docx-wrapper').style.display !== 'none') {
    applyDocxZoom(scale);
  } else if ($('pdf-wrapper').style.display !== 'none') {
    applyPdfZoom(scale);
  }
}

function resetZoom() {
  isAutoFit = false;
  if ($('docx-wrapper').style.display !== 'none') {
    applyDocxZoom(1.0);
  } else if ($('pdf-wrapper').style.display !== 'none') {
    applyPdfZoom(1.0);
  }
}

function fitCurrentToWidth() {
  isAutoFit = true;
  if ($('docx-wrapper').style.display !== 'none') {
    fitDocxToWidth();
  } else if ($('pdf-wrapper').style.display !== 'none') {
    fitPdfToWidth();
  }
}

// ---------------------------------------------------------------------------
// Native System Actions
// ---------------------------------------------------------------------------
async function openInDefaultApp() {
  if (!currentFile || !currentFile.path || currentFile.isMemoryOnly) return;
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
  if (!currentFile || !currentFile.path || currentFile.isMemoryOnly) return;
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

async function copyToClipboard(text) {
  // 1. Modern Clipboard API first
  if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // Fallback if iframe Permissions-Policy denies Clipboard API
    }
  }

  // 2. Legacy execCommand fallback
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.style.top = '0';
    ta.setAttribute('readonly', '');
    document.body.appendChild(ta);
    ta.select();
    ta.setSelectionRange(0, text.length);
    const copied = document.execCommand('copy');
    document.body.removeChild(ta);
    if (copied) return true;
  } catch {
    // Fallback to OS command
  }

  // 3. Sidecar backend OS clipboard API
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
  if (!currentFile || !currentFile.path || currentFile.isMemoryOnly) return;
  const ok = await copyToClipboard(currentFile.path);
  if (ok) {
    toast('文件绝对路径已复制到剪贴板', 'success');
  } else {
    toast('复制失败，请检查剪贴板权限', 'error');
  }
}

// ---------------------------------------------------------------------------
// Drag and Drop
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
      const ext = name.includes('.') ? '.' + name.split('.').pop().toLowerCase() : '';
      const localPath = file.path || '';
      markUserActivity();

      if (localPath) {
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
  const token = ++activeRenderToken;
  cleanupPdfState();

  const ext = file.name.includes('.') ? '.' + file.name.split('.').pop().toLowerCase() : '';
  currentFile = {
    name: file.name,
    path: '',
    size: file.size,
    mtime: file.lastModified,
    ext,
    isMemoryOnly: true,
  };
  updateButtonsState(currentFile);
  updateMetaDisplay(currentFile);
  $('header-status').textContent = '拖放预览';
  showLoading(`正在读取本地文件 ${file.name}...`);

  try {
    const arrayBuffer = await file.arrayBuffer();
    if (token !== activeRenderToken) return;

    if (ext === '.docx') {
      await mountDocxBuffer(arrayBuffer, token);
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      await ensureXlsxEngine();
      if (token !== activeRenderToken) return;

      const wb = ext === '.csv'
        ? window.XLSX.read(decodeTextBuffer(arrayBuffer), { type: 'string' })
        : window.XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

      const { sheetNames, sheetsMap } = extractSheetsMapFromWorkbook(wb);
      mountSpreadsheetView(sheetNames, sheetsMap);
    } else if (ext === '.pdf') {
      await mountPdfBuffer(arrayBuffer, token);
    } else if (ext === '.txt' || ext === '.md' || ext === '.json') {
      mountTextContent(decodeTextBuffer(arrayBuffer), ext);
    } else {
      showEmptyState(`暂不支持直接脱机解析 ${ext || '该'} 文件`);
    }
  } catch (err) {
    if (token !== activeRenderToken) return;
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

  const reloadUiBtn = $('btn-reload-ui');
  if (reloadUiBtn) {
    reloadUiBtn.addEventListener('click', async () => {
      markUserActivity();
      await reloadCurrentDocument(true);
    });
  }

  $('btn-zoom-in').addEventListener('click', () => updateZoom(currentZoom + 10));
  $('btn-zoom-out').addEventListener('click', () => updateZoom(currentZoom - 10));
  $('btn-zoom-reset').addEventListener('click', resetZoom);
  $('btn-zoom-fit').addEventListener('click', fitCurrentToWidth);

  // Ctrl/Cmd + Wheel zoom for DOCX and PDF (plain wheel triggers zero network calls)
  $('viewport-container').addEventListener('wheel', (e) => {
    if (e.ctrlKey || e.metaKey) {
      const isDocx = $('docx-wrapper').style.display !== 'none';
      const isPdf = $('pdf-wrapper').style.display !== 'none';
      if (isDocx || isPdf) {
        e.preventDefault();
        const step = e.deltaY < 0 ? 10 : -10;
        updateZoom(currentZoom + step);
      }
    }
  }, { passive: false });

  setupDragAndDrop();

  if (window.ResizeObserver) {
    let resizeRaf = null;
    const ro = new ResizeObserver(() => {
      if (!isAutoFit || !currentFile) return;
      if (resizeRaf) cancelAnimationFrame(resizeRaf);
      resizeRaf = requestAnimationFrame(() => {
        if (!isAutoFit || !currentFile) return;
        if (currentFile.ext === '.docx') {
          fitDocxToWidth();
        } else if (currentFile.ext === '.pdf') {
          fitPdfToWidth();
        }
      });
    });
    ro.observe($('viewport-container'));
  }

  window.addEventListener('popstate', () => syncWithUrl(true));
  window.addEventListener('hashchange', () => syncWithUrl(true));

  window.addEventListener('focus', () => {
    markUserActivity();
    checkAndReloadIfModified(true);
  });

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) {
      markUserActivity();
      checkAndReloadIfModified(true);
    }
  });

  // In-memory URL search watcher (zero network or disk I/O)
  setInterval(() => {
    const currentSearch = window.location.search || '';
    if (currentSearch !== lastKnownSearch) {
      lastKnownSearch = currentSearch;
      syncWithUrl(true);
    }
  }, 250);

  startAdaptiveHeartbeat();
}

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  initEvents();
  setTimeout(() => {
    lastKnownSearch = window.location.search || '';
    syncWithUrl(true);
  }, 50);
});
