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
let currentWorkbook = null;
let activeSheetName = null;
let toastTimer = null;
let lastKnownQueryFile = null;

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

function getUrlFileParam() {
  try {
    const params = new URLSearchParams(window.location.search);
    return params.get('file') || params.get('path') || null;
  } catch {
    return null;
  }
}

// Normalize path separators for comparison
function normPath(p) {
  return (p || '').replace(/\\/g, '/').toLowerCase();
}

// ---------------------------------------------------------------------------
// Sniffer & File Discovery
// ---------------------------------------------------------------------------
async function refreshFiles(preferPath = null) {
  try {
    let workspaceUris = [];
    if (sidecar.getWorkspaceUris) {
      workspaceUris = await sidecar.getWorkspaceUris().catch(() => []);
    }

    const conversationId = sidecar.conversationId || '';
    const params = new URLSearchParams();
    if (conversationId) params.set('conversationId', conversationId);
    if (workspaceUris.length > 0) params.set('workspaces', JSON.stringify(workspaceUris));

    const res = await callApi(`/api/files/sniff?${params.toString()}`);
    const newFiles = res.files || [];
    sniffedFiles = newFiles;

    const urlFile = getUrlFileParam();
    const effectivePrefer = preferPath || urlFile;

    renderFileDropdown(effectivePrefer);
    $('header-status').textContent = `已发现 ${sniffedFiles.length} 个文档`;
  } catch (err) {
    console.error('[sniff] Error:', err);
    $('header-status').textContent = '嗅探出错';
  }
}

function renderFileDropdown(preferPath = null) {
  const select = $('file-select');
  select.innerHTML = '';

  if (sniffedFiles.length === 0) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '暂未嗅探到 Office 文档';
    select.appendChild(opt);
    updateButtonsState(null);
    showEmptyState('未发现任何 Office 文档，请在上方手动输入路径或拖入文件');
    return;
  }

  // Group files by source: conversation vs workspace vs recent
  const groups = {
    conversation: { label: '💬 当前会话工件', items: [] },
    workspace: { label: '📁 当前工作区文档', items: [] },
    recent: { label: '🕒 最近打开历史', items: [] },
  };

  for (const f of sniffedFiles) {
    const grp = groups[f.source] || groups.workspace;
    grp.items.push(f);
  }

  let targetToSelect = null;

  // 1. If explicit path requested (e.g. from ?file=... or direct user select)
  if (preferPath) {
    const normPrefer = normPath(preferPath);
    targetToSelect = sniffedFiles.find((f) => normPath(f.path) === normPrefer);
    // If not found in sniffed list, synthesize a custom item
    if (!targetToSelect) {
      const name = preferPath.split(/[\/\\]/).pop();
      const ext = '.' + name.split('.').pop().toLowerCase();
      targetToSelect = {
        name,
        path: preferPath,
        ext,
        size: 0,
        mtime: Date.now(),
        source: 'direct',
      };
      sniffedFiles.unshift(targetToSelect);
    }
  }

  // 2. If no explicit path, check if newest file is newer than currentFile
  if (!targetToSelect) {
    const newestFile = sniffedFiles[0];
    if (!currentFile) {
      targetToSelect = newestFile;
    } else {
      const currentStillExists = sniffedFiles.some((f) => normPath(f.path) === normPath(currentFile.path));
      // Auto-switch if newest file was updated after currentFile
      if (newestFile && newestFile.mtime > (currentFile.mtime || 0) + 1500) {
        targetToSelect = newestFile;
      } else if (currentStillExists) {
        targetToSelect = sniffedFiles.find((f) => normPath(f.path) === normPath(currentFile.path));
      } else {
        targetToSelect = newestFile;
      }
    }
  }

  // Build optgroups
  for (const [key, grp] of Object.entries(groups)) {
    if (grp.items.length === 0) continue;
    const optgroup = document.createElement('optgroup');
    optgroup.label = grp.label;
    for (const f of grp.items) {
      const opt = document.createElement('option');
      opt.value = f.path;
      const icon = getFileIcon(f.ext);
      opt.textContent = `${icon} ${f.name} (${formatBytes(f.size)})`;
      optgroup.appendChild(opt);
    }
    select.appendChild(optgroup);
  }

  if (targetToSelect) {
    select.value = targetToSelect.path;
    // Only reload if file actually changed
    if (!currentFile || normPath(currentFile.path) !== normPath(targetToSelect.path)) {
      loadFile(targetToSelect);
    }
  }
}

function getFileIcon(ext) {
  switch ((ext || '').toLowerCase()) {
    case '.docx':
    case '.doc':
      return '📄';
    case '.xlsx':
    case '.xls':
    case '.csv':
      return '📊';
    case '.pdf':
      return '📕';
    case '.pptx':
    case '.ppt':
      return '📑';
    default:
      return '📝';
  }
}

// ---------------------------------------------------------------------------
// Load and Render File
// ---------------------------------------------------------------------------
async function loadFile(fileInfo) {
  if (!fileInfo || !fileInfo.path) return;
  currentFile = fileInfo;
  updateButtonsState(fileInfo);
  updateMetaDisplay(fileInfo);

  showLoading(`正在载入 ${fileInfo.name}...`);

  const ext = (fileInfo.ext || '').toLowerCase();
  try {
    if (ext === '.docx') {
      await renderDocx(fileInfo.path);
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      await renderSpreadsheet(fileInfo.path, ext);
    } else if (ext === '.pdf') {
      renderPdf(fileInfo.path);
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
  $('btn-ask-agent').disabled = !hasFile;

  if (file) {
    const ext = (file.ext || '').toLowerCase();
    let label = '在系统默认程序中打开';
    if (ext === '.docx' || ext === '.doc') label = '在 Word / WPS 中打开';
    else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') label = '在 Excel / WPS 表格中打开';
    else if (ext === '.pdf') label = '在 PDF 阅读器中打开';
    $('btn-open-label').textContent = label;
  } else {
    $('btn-open-label').textContent = '在系统默认程序中打开';
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

// 1. DOCX Renderer (docx-preview)
async function renderDocx(filePath) {
  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}`);
  if (!res.ok) throw new Error('无法读取文件内容');
  const arrayBuffer = await res.arrayBuffer();

  hideAllViewports();
  $('docx-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'flex';

  const container = $('docx-render-target');
  container.innerHTML = '';

  resetZoom();

  if (window.docx && window.docx.renderAsync) {
    await window.docx.renderAsync(arrayBuffer, container, null, {
      inWrapper: false,
      ignoreWidth: false,
      ignoreHeight: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
    });
  } else {
    throw new Error('docx-preview 引擎未就绪');
  }
}

// 2. Spreadsheet Renderer (SheetJS / XLSX)
async function renderSpreadsheet(filePath, ext) {
  const res = await sidecar.fetch(`/api/file/raw?path=${encodeURIComponent(filePath)}`);
  if (!res.ok) throw new Error('无法读取表格内容');
  const arrayBuffer = await res.arrayBuffer();

  hideAllViewports();
  $('xlsx-wrapper').style.display = 'flex';
  $('zoom-controls').style.display = 'none';

  if (!window.XLSX) {
    throw new Error('XLSX 引擎未就绪');
  }

  const data = new Uint8Array(arrayBuffer);
  currentWorkbook = window.XLSX.read(data, { type: 'array' });
  const sheetNames = currentWorkbook.SheetNames || [];

  if (sheetNames.length === 0) {
    $('xlsx-table-wrapper').innerHTML = '<div style="padding:20px;text-align:center;">工作簿为空</div>';
    return;
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
}

function renderSheetTable(sheetName) {
  activeSheetName = sheetName;
  const worksheet = currentWorkbook.Sheets[sheetName];
  if (!worksheet) return;

  const html = window.XLSX.utils.sheet_to_html(worksheet, {
    header: '',
    footer: '',
    id: 'sheet-table',
    editable: false,
  });

  const tableContainer = $('xlsx-table-wrapper');
  tableContainer.innerHTML = html;
  const table = tableContainer.querySelector('table');
  if (table) {
    table.className = 'sheet-table';
  }
}

// 3. PDF Renderer
function renderPdf(filePath) {
  hideAllViewports();
  const iframe = $('pdf-iframe');
  iframe.style.display = 'block';
  $('zoom-controls').style.display = 'none';
  iframe.src = `/api/file/raw?path=${encodeURIComponent(filePath)}`;
}

// ---------------------------------------------------------------------------
// Viewport State Switches
// ---------------------------------------------------------------------------
function hideAllViewports() {
  $('empty-state').style.display = 'none';
  $('loading-state').style.display = 'none';
  $('docx-wrapper').style.display = 'none';
  $('xlsx-wrapper').style.display = 'none';
  $('pdf-iframe').style.display = 'none';
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
// Zoom Controls (DOCX)
// ---------------------------------------------------------------------------
function updateZoom(newZoom) {
  currentZoom = Math.max(50, Math.min(200, newZoom));
  $('zoom-label').textContent = `${currentZoom}%`;
  const wrapper = $('docx-wrapper');
  wrapper.style.transform = `scale(${currentZoom / 100})`;
  wrapper.style.transformOrigin = 'top center';
}

function resetZoom() {
  currentZoom = 100;
  $('zoom-label').textContent = '100%';
  const wrapper = $('docx-wrapper');
  wrapper.style.transform = 'scale(1)';
}

// ---------------------------------------------------------------------------
// Native System Actions (P0)
// ---------------------------------------------------------------------------
async function openInDefaultApp() {
  if (!currentFile) return;
  const btn = $('btn-open-default');
  btn.disabled = true;
  try {
    const res = await callApi('/api/action/open-default', {
      method: 'POST',
      body: { path: currentFile.path },
    });
    if (res.ok) {
      toast(res.message || '已成功在默认程序中打开', 'success');
    } else {
      toast(res.error || '打开失败', 'error');
    }
  } catch (err) {
    toast(`启动失败: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
  }
}

async function revealInExplorer() {
  if (!currentFile) return;
  try {
    const res = await callApi('/api/action/reveal', {
      method: 'POST',
      body: { path: currentFile.path },
    });
    if (res.ok) {
      toast(res.message || '已在资源管理器中高亮定位', 'success');
    } else {
      toast(res.error || '定位失败', 'error');
    }
  } catch (err) {
    toast(`定位失败: ${err.message}`, 'error');
  }
}

function copyFilePath() {
  if (!currentFile) return;
  navigator.clipboard.writeText(currentFile.path).then(
    () => toast('文件绝对路径已复制到剪贴板', 'success'),
    (err) => toast(`复制失败: ${err.message}`, 'error')
  );
}

async function askAgentToPolish() {
  if (!currentFile) return;
  if (!sidecar.agent || !sidecar.agent.sendMessage) {
    toast('宿主 Agent 接口不可用', 'error');
    return;
  }
  const msg = `请针对我正在预览的文档 [${currentFile.name}]（路径：${currentFile.path}）进行审查与进一步优化润色：\n`;
  try {
    await sidecar.agent.sendMessage(msg);
    toast('已将文档润色请求推送至当前会话', 'success');
  } catch (err) {
    toast(`推送失败: ${err.message}`, 'error');
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
      resetZoom();
      await window.docx.renderAsync(arrayBuffer, $('docx-render-target'), null, { inWrapper: false });
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      hideAllViewports();
      $('xlsx-wrapper').style.display = 'flex';
      $('zoom-controls').style.display = 'none';
      const data = new Uint8Array(arrayBuffer);
      currentWorkbook = window.XLSX.read(data, { type: 'array' });
      renderSheetTable(currentWorkbook.SheetNames[0]);
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
  $('btn-ask-agent').addEventListener('click', askAgentToPolish);

  $('file-select').addEventListener('change', (e) => {
    const selectedPath = e.target.value;
    const found = sniffedFiles.find((f) => f.path === selectedPath);
    if (found) loadFile(found);
  });

  $('btn-refresh').addEventListener('click', () => refreshFiles());

  $('btn-toggle-manual').addEventListener('click', () => {
    const bar = $('manual-path-bar');
    bar.style.display = bar.style.display === 'none' ? 'flex' : 'none';
    if (bar.style.display === 'flex') $('input-manual-path').focus();
  });

  $('btn-load-manual').addEventListener('click', () => {
    const inputPath = $('input-manual-path').value.trim();
    if (!inputPath) return;
    const name = inputPath.split(/[\/\\]/).pop();
    const ext = '.' + name.split('.').pop().toLowerCase();
    loadFile({
      name,
      path: inputPath,
      size: 0,
      mtime: Date.now(),
      ext,
      source: 'manual',
    });
  });

  $('input-manual-path').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') $('btn-load-manual').click();
  });

  $('btn-zoom-in').addEventListener('click', () => updateZoom(currentZoom + 15));
  $('btn-zoom-out').addEventListener('click', () => updateZoom(currentZoom - 15));
  $('btn-zoom-reset').addEventListener('click', resetZoom);

  setupDragAndDrop();

  // Focus & Visibility Auto-Refresh
  window.addEventListener('focus', () => refreshFiles());
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) refreshFiles();
  });

  // URL query parameter watcher
  window.addEventListener('popstate', () => refreshFiles());
  window.addEventListener('hashchange', () => refreshFiles());

  // Periodically check for URL search param changes (e.g. when host re-triggers sidecar view)
  setInterval(() => {
    const currentQueryFile = getUrlFileParam();
    if (currentQueryFile && currentQueryFile !== lastKnownQueryFile) {
      lastKnownQueryFile = currentQueryFile;
      refreshFiles(currentQueryFile);
    }
  }, 1000);

  if (sidecar.onWorkspaceChange) {
    sidecar.onWorkspaceChange(() => {
      refreshFiles();
    });
  }
}

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  initEvents();
  setTimeout(() => {
    lastKnownQueryFile = getUrlFileParam();
    refreshFiles(lastKnownQueryFile);
  }, 100);
});
