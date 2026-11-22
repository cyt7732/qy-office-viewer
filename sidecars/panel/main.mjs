// QY Office Viewer Backend
// Sidecar Node.js Service for Antigravity 2.22.0+

import { createContext, runInContext } from 'node:vm';
import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { homedir, platform, release } from 'node:os';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { exec, execFile, spawn } from 'node:child_process';
import { Response, SidecarApp } from 'sidecar_sdk';

const HERE = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = process.env.ANTIGRAVITY_EXECUTABLE_DATA_DIR || join(HERE, '.data');
const HISTORY_FILE = join(DATA_DIR, 'recent_files.json');
const STARTED_AT = Date.now();

const app = new SidecarApp();

// Helper to read local files in UTF-8
const readLocal = (name) => readFileSync(join(HERE, name), 'utf8');

// ---------------------------------------------------------------------------
// Static Routes (Strict No-Cache for Instant Hot Reloading)
// ---------------------------------------------------------------------------
app.page('/', () =>
  new Response(readLocal('index.html'), {
    contentType: 'text/html; charset=utf-8',
    headers: { 'Cache-Control': 'no-store, no-cache, must-revalidate', 'Pragma': 'no-cache' },
  })
);

app.api('/app.js', () =>
  new Response(readLocal('app.js'), {
    contentType: 'application/javascript',
    headers: { 'Cache-Control': 'no-store, no-cache, must-revalidate', 'Pragma': 'no-cache' },
  }),
'GET');

app.api('/styles.css', () =>
  new Response(readLocal('styles.css'), {
    contentType: 'text/css',
    headers: { 'Cache-Control': 'no-store, no-cache, must-revalidate', 'Pragma': 'no-cache' },
  }),
'GET');

// Vendor bundles
app.api('/vendor/docx-preview.min.js', () =>
  new Response(readLocal('vendor/docx-preview.min.js'), { contentType: 'application/javascript' }),
'GET');

app.api('/vendor/jszip.min.js', () =>
  new Response(readLocal('vendor/jszip.min.js'), { contentType: 'application/javascript' }),
'GET');

app.api('/vendor/xlsx.mini.min.js', () =>
  new Response(readLocal('vendor/xlsx.full.min.js'), { contentType: 'application/javascript' }),
'GET');

app.api('/vendor/xlsx.full.min.js', () =>
  new Response(readLocal('vendor/xlsx.full.min.js'), { contentType: 'application/javascript' }),
'GET');

app.api('/vendor/pdf.min.js', () =>
  new Response(readLocal('vendor/pdf.min.js'), {
    contentType: 'application/javascript; charset=utf-8',
    headers: { 'Cache-Control': 'public, max-age=3600' },
  }),
'GET');

app.api('/vendor/pdf.worker.min.js', () =>
  new Response(readLocal('vendor/pdf.worker.min.js'), {
    contentType: 'application/javascript; charset=utf-8',
    headers: { 'Cache-Control': 'public, max-age=3600' },
  }),
'GET');

// ---------------------------------------------------------------------------
// System & Diagnostic APIs
// ---------------------------------------------------------------------------
app.api('/api/status', () => ({
  name: 'qy-office-viewer',
  version: '1.1.0',
  node: process.version,
  platform: `${platform()} ${release()}`,
  uptimeSeconds: Math.round((Date.now() - STARTED_AT) / 1000),
}), 'GET');

// ---------------------------------------------------------------------------
// File Sniffer API (Strictly Office Documents Only)
// ---------------------------------------------------------------------------
const OFFICE_EXTS = new Set([
  '.docx', '.doc', '.xlsx', '.xls', '.csv', '.pdf', '.pptx', '.ppt',
]);

function getFileType(ext) {
  switch (ext.toLowerCase()) {
    case '.docx':
    case '.doc':
      return 'word';
    case '.xlsx':
    case '.xls':
    case '.csv':
      return 'excel';
    case '.pdf':
      return 'pdf';
    case '.pptx':
    case '.ppt':
      return 'ppt';
    default:
      return 'other';
  }
}

function safeStat(filePath) {
  try {
    const st = statSync(filePath);
    return {
      size: st.size,
      mtime: st.mtimeMs,
      isFile: st.isFile(),
      isDir: st.isDirectory(),
    };
  } catch {
    return null;
  }
}

function resolveActualPath(inputPath) {
  if (!inputPath || typeof inputPath !== 'string') return null;
  const targetPath = inputPath.trim().replace(/^["']|["']$/g, '');
  if (existsSync(targetPath)) return targetPath;

  try {
    const dir = dirname(targetPath);
    if (!existsSync(dir)) return null;

    const targetBase = basename(targetPath);
    const targetExt = extname(targetBase).toLowerCase();
    const targetStem = basename(targetBase, extname(targetBase)).toLowerCase();

    const entries = readdirSync(dir, { withFileTypes: true });
    const candidates = [];

    for (const ent of entries) {
      if (!ent.isFile()) continue;
      const fName = ent.name;
      const fExt = extname(fName).toLowerCase();
      if (fExt !== targetExt) continue;

      const fStem = basename(fName, extname(fName)).toLowerCase();
      if (fStem === targetStem) {
        return join(dir, fName);
      }

      const cleanTarget = targetStem.replace(/[\s_\-\(\)（）【】\[\]]/g, '');
      const cleanCandidate = fStem.replace(/[\s_\-\(\)（）【】\[\]]/g, '');

      let matches = 0;
      const candidateChars = Array.from(cleanCandidate);
      for (const ch of Array.from(cleanTarget)) {
        const idx = candidateChars.indexOf(ch);
        if (idx !== -1) {
          matches++;
          candidateChars.splice(idx, 1);
        }
      }
      const similarity = (2 * matches) / (cleanTarget.length + cleanCandidate.length);
      candidates.push({ name: fName, path: join(dir, fName), similarity });
    }

    if (candidates.length === 1 && candidates[0].similarity >= 0.5) {
      console.log(`[resolveActualPath] 唯一候选文件智能对齐: "${targetBase}" -> "${candidates[0].name}" (相似度: ${(candidates[0].similarity * 100).toFixed(1)}%)`);
      return candidates[0].path;
    }

    if (candidates.length > 1) {
      candidates.sort((a, b) => b.similarity - a.similarity);
      if (candidates[0].similarity >= 0.75) {
        console.log(`[resolveActualPath] 高相似度候选文件智能对齐: "${targetBase}" -> "${candidates[0].name}" (相似度: ${(candidates[0].similarity * 100).toFixed(1)}%)`);
        return candidates[0].path;
      }
    }
  } catch (err) {
    console.warn('[resolveActualPath] Notice:', err.message);
  }

  return null;
}


function scanDirForFiles(dirPath, maxDepth = 2, currentDepth = 0, ignoreDirs = new Set(['.git', 'node_modules', '.venv', 'dist', 'build', '.data', '.system_generated', '.user_uploaded', 'scratch'])) {
  if (currentDepth > maxDepth || !existsSync(dirPath)) return [];
  const results = [];
  try {
    const entries = readdirSync(dirPath, { withFileTypes: true });
    for (const ent of entries) {
      if (ent.name.startsWith('.') && ent.name !== '.gemini') {
        if (ent.isDirectory()) continue;
      }
      const fullPath = join(dirPath, ent.name);
      if (ent.isDirectory()) {
        if (!ignoreDirs.has(ent.name.toLowerCase())) {
          results.push(...scanDirForFiles(fullPath, maxDepth, currentDepth + 1, ignoreDirs));
        }
      } else if (ent.isFile()) {
        const ext = extname(ent.name).toLowerCase();
        if (OFFICE_EXTS.has(ext)) {
          const st = safeStat(fullPath);
          if (st) {
            results.push({
              name: ent.name,
              path: fullPath,
              size: st.size,
              mtime: st.mtime,
              ext,
              type: getFileType(ext),
            });
          }
        }
      }
    }
  } catch (err) {
    console.error(`[sniff] Error reading dir ${dirPath}:`, err.message);
  }
  return results;
}

// GET /api/files/sniff: Discovers documents from conversation artifacts, active workspaces & recent history
app.api('/api/files/sniff', (data) => {
  const conversationId = data.conversationId;
  let workspacePaths = [];
  if (data.workspaces) {
    try {
      const parsed = typeof data.workspaces === 'string' ? JSON.parse(data.workspaces) : data.workspaces;
      if (Array.isArray(parsed)) {
        workspacePaths = parsed.map((u) => u.startsWith('file://') ? fileURLToPath(u) : u);
      }
    } catch {
      // ignore JSON parse error
    }
  }

  const fileMap = new Map();

  // 1. Sniff conversation artifacts directory
  if (conversationId) {
    const convDir = join(homedir(), '.gemini', 'antigravity', 'brain', conversationId);
    if (existsSync(convDir)) {
      const convFiles = scanDirForFiles(convDir, 1);
      for (const item of convFiles) {
        fileMap.set(item.path, { ...item, source: 'conversation' });
      }
    }
  }

  // 2. Sniff workspaces
  for (const wsPath of workspacePaths) {
    if (existsSync(wsPath)) {
      const wsFiles = scanDirForFiles(wsPath, 2);
      for (const item of wsFiles) {
        if (!fileMap.has(item.path)) {
          fileMap.set(item.path, { ...item, source: 'workspace' });
        }
      }
    }
  }

  // 3. Include recent history files (only if they are valid office files)
  const recent = loadRecentFiles();
  for (const item of recent) {
    if (existsSync(item.path) && !fileMap.has(item.path)) {
      const ext = extname(item.path).toLowerCase();
      if (OFFICE_EXTS.has(ext)) {
        const st = safeStat(item.path);
        if (st) {
          fileMap.set(item.path, {
            name: basename(item.path),
            path: item.path,
            size: st.size,
            mtime: st.mtime,
            ext,
            type: getFileType(ext),
            source: 'recent',
          });
        }
      }
    }
  }

  // Convert map to array and sort by modified time desc (newest first)
  const allFiles = Array.from(fileMap.values()).sort((a, b) => b.mtime - a.mtime);
  return { files: allFiles };
}, 'GET');

// ---------------------------------------------------------------------------
// File Content & Meta APIs
// ---------------------------------------------------------------------------
const MIME_TYPES = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pdf': 'application/pdf',
  '.doc': 'application/msword',
  '.xls': 'application/vnd.ms-excel',
  '.csv': 'text/csv; charset=utf-8',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.ppt': 'application/vnd.ms-powerpoint',
};

// GET /api/file/raw?path=...
app.api('/api/file/raw', (data) => {
  const actualPath = resolveActualPath(data.path);
  if (!actualPath) {
    return new Response(JSON.stringify({ error: '文件不存在或路径无效' }), {
      status: 404,
      contentType: 'application/json',
    });
  }

  const ext = extname(actualPath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  try {
    const buffer = readFileSync(actualPath);
    // Record to recent history
    addRecentFile(actualPath);
    return new Response(buffer, {
      contentType,
      headers: {
        'Content-Disposition': `inline; filename="${encodeURIComponent(basename(actualPath))}"`,
        'Cache-Control': 'no-store, no-cache, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0',
      },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      contentType: 'application/json',
    });
  }
}, 'GET');

// GET /api/file/meta?path=...
app.api('/api/file/meta', (data) => {
  const actualPath = resolveActualPath(data.path);
  if (!actualPath) {
    return { error: '文件不存在或路径无效' };
  }
  const st = safeStat(actualPath);
  const ext = extname(actualPath).toLowerCase();
  return {
    name: basename(actualPath),
    path: actualPath,
    originalPath: data.path,
    resolved: actualPath !== data.path,
    size: st?.size ?? 0,
    mtime: st?.mtime ?? 0,
    ext,
    type: getFileType(ext),
  };
}, 'GET');

// Server-side Spreadsheet Parser (100% CJK & CP936 Encoding Lossless Guarantee)
let cachedXLSX = null;
function getBackendXLSX() {
  if (!cachedXLSX) {
    const sandbox = { window: {}, Uint8Array, ArrayBuffer, Buffer, console };
    createContext(sandbox);
    const code = readFileSync(join(HERE, 'vendor/xlsx.full.min.js'), 'utf8');
    runInContext(code, sandbox);
    cachedXLSX = sandbox.window.XLSX || sandbox.XLSX;
  }
  return cachedXLSX;
}

// GET /api/file/spreadsheet?path=...
app.api('/api/file/spreadsheet', (data) => {
  const actualPath = resolveActualPath(data.path);
  if (!actualPath) {
    return { ok: false, error: '文件不存在或路径无效' };
  }

  try {
    const xlsxLib = getBackendXLSX();
    const buf = readFileSync(actualPath);
    const wb = xlsxLib.read(buf, { type: 'buffer' });
    const sheetNames = wb.SheetNames || [];
    const sheets = {};

    for (const name of sheetNames) {
      const ws = wb.Sheets[name];
      if (ws) {
        sheets[name] = xlsxLib.utils.sheet_to_html(ws, {
          header: '',
          footer: '',
          id: 'sheet-table',
          editable: false,
        });
      }
    }

    addRecentFile(actualPath);
    return {
      ok: true,
      sheetNames,
      sheets,
    };
  } catch (err) {
    console.error('[spreadsheet] Backend parse error:', err);
    return { ok: false, error: err.message };
  }
}, 'GET');

// ---------------------------------------------------------------------------
// Native System Actions (P0)
// ---------------------------------------------------------------------------

// POST /api/action/open-default { path }
app.api('/api/action/open-default', (data) => {
  const actualPath = resolveActualPath(data.path);
  if (!actualPath) {
    return { ok: false, error: '文件不存在或路径无效' };
  }

  const absPath = resolve(actualPath);
  addRecentFile(absPath);

  const currentPlatform = platform();
  try {
    if (currentPlatform === 'win32') {
      const child = spawn('cmd.exe', ['/c', 'start', '', absPath], {
        detached: true,
        stdio: 'ignore',
        windowsHide: true,
      });
      child.unref();
    } else if (currentPlatform === 'darwin') {
      const child = spawn('open', [absPath], { detached: true, stdio: 'ignore' });
      child.unref();
    } else {
      const child = spawn('xdg-open', [absPath], { detached: true, stdio: 'ignore' });
      child.unref();
    }
    return { ok: true, message: `已在系统默认程序中打开：${basename(absPath)}` };
  } catch (err) {
    console.error('[open-default] Failed:', err);
    return { ok: false, error: err.message };
  }
});

// POST /api/action/reveal { path }
app.api('/api/action/reveal', (data) => {
  const actualPath = resolveActualPath(data.path);
  if (!actualPath) {
    return { ok: false, error: '文件不存在或路径无效' };
  }

  const absPath = resolve(actualPath);
  const currentPlatform = platform();

  try {
    if (currentPlatform === 'win32') {
      const child = spawn('explorer.exe', [`/select,${absPath}`], {
        detached: true,
        stdio: 'ignore',
      });
      child.unref();
    } else if (currentPlatform === 'darwin') {
      const child = spawn('open', ['-R', absPath], { detached: true, stdio: 'ignore' });
      child.unref();
    } else {
      const child = spawn('xdg-open', [dirname(absPath)], { detached: true, stdio: 'ignore' });
      child.unref();
    }
    return { ok: true, message: `已在资源管理器中高亮定位：${basename(absPath)}` };
  } catch (err) {
    console.error('[reveal] Failed:', err);
    return { ok: false, error: err.message };
  }
});

// POST /api/action/copy-clipboard { text }
app.api('/api/action/copy-clipboard', (data) => {
  const text = data.text;
  if (!text) {
    return { ok: false, error: '文本内容为空' };
  }

  return new Promise((resolveResult) => {
    const currentPlatform = platform();
    if (currentPlatform === 'win32') {
      const b64 = Buffer.from(text, 'utf8').toString('base64');
      execFile('powershell.exe', ['-NoProfile', '-Command', `Set-Clipboard -Value ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('${b64}')))`], (err) => {
        if (err) {
          console.error('[copy-clipboard] Failed:', err);
          resolveResult({ ok: false, error: err.message });
        } else {
          resolveResult({ ok: true, message: '已成功复制到系统剪贴板' });
        }
      });
    } else if (currentPlatform === 'darwin') {
      const proc = execFile('pbcopy', [], (err) => {
        if (err) resolveResult({ ok: false, error: err.message });
        else resolveResult({ ok: true, message: '已复制到剪贴板' });
      });
      proc.stdin.write(text);
      proc.stdin.end();
    } else {
      const proc = execFile('xclip', ['-selection', 'clipboard'], (err) => {
        if (err) resolveResult({ ok: false, error: err.message });
        else resolveResult({ ok: true, message: '已复制到剪贴板' });
      });
      proc.stdin.write(text);
      proc.stdin.end();
    }
  });
});

// ---------------------------------------------------------------------------
// Recent Files Persistence
// ---------------------------------------------------------------------------
function loadRecentFiles() {
  if (!existsSync(HISTORY_FILE)) return [];
  try {
    return JSON.parse(readFileSync(HISTORY_FILE, 'utf8'));
  } catch {
    return [];
  }
}

function addRecentFile(filePath) {
  try {
    mkdirSync(DATA_DIR, { recursive: true });
    let list = loadRecentFiles().filter((i) => i.path !== filePath);
    list.unshift({ path: filePath, openedAt: Date.now() });
    if (list.length > 20) list = list.slice(0, 20);
    writeFileSync(HISTORY_FILE, JSON.stringify(list, null, 2), 'utf8');
  } catch (err) {
    console.error('[recent] Failed to save recent file:', err.message);
  }
}

app.run();
