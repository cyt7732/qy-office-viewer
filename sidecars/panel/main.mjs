// QY Office Viewer Backend
// Sidecar Node.js Service for Antigravity

import { createContext, runInContext } from 'node:vm';
import { existsSync, readFileSync, watch } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { platform, release } from 'node:os';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFile, spawn } from 'node:child_process';
import { Response, SidecarApp } from 'sidecar_sdk';
import {
  MIME_TYPES,
  buildWinClipboardArgs,
  buildWinOpenArgs,
  buildWinRevealArgs,
  decodeCsvBuffer,
  findSimilarCandidates,
  getFileType,
  isAllowedPreviewExt,
  isSpreadsheetExt,
  resolveExactPath,
  resolveWinFallbackPowerShellExe,
  safeStat,
  sanitizeCmapFileName,
} from './lib.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const STARTED_AT = Date.now();
const app = new SidecarApp();

const NO_CACHE_HEADERS = {
  'Cache-Control': 'no-store, no-cache, must-revalidate',
  'Pragma': 'no-cache',
};

const STATIC_CACHE_HEADERS = {
  'Cache-Control': 'public, max-age=86400',
};

const readLocalText = (name) => readFileSync(join(HERE, name), 'utf8');

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    contentType: 'application/json; charset=utf-8',
  });
}

// ---------------------------------------------------------------------------
// Static & Vendor Routes
// ---------------------------------------------------------------------------
app.page('/', () =>
  new Response(readLocalText('index.html'), {
    contentType: 'text/html; charset=utf-8',
    headers: NO_CACHE_HEADERS,
  })
);

app.api('/app.js', () =>
  new Response(readLocalText('app.js'), {
    contentType: 'application/javascript; charset=utf-8',
    headers: NO_CACHE_HEADERS,
  }),
'GET');

app.api('/lib.mjs', () =>
  new Response(readLocalText('lib.mjs'), {
    contentType: 'application/javascript; charset=utf-8',
    headers: NO_CACHE_HEADERS,
  }),
'GET');

app.api('/styles.css', () =>
  new Response(readLocalText('styles.css'), {
    contentType: 'text/css; charset=utf-8',
    headers: NO_CACHE_HEADERS,
  }),
'GET');

// In-memory cached vendor bundles (zero repeated disk I/O)
const vendorCache = new Map();
const VENDOR_ASSETS = new Map([
  ['/vendor/docx-preview.min.js', 'vendor/docx-preview.min.js'],
  ['/vendor/jszip.min.js', 'vendor/jszip.min.js'],
  ['/vendor/xlsx.full.min.js', 'vendor/xlsx.full.min.js'],
  ['/vendor/pdf.min.js', 'vendor/pdf.min.js'],
  ['/vendor/pdf.worker.min.js', 'vendor/pdf.worker.min.js'],
]);

for (const [route, relPath] of VENDOR_ASSETS.entries()) {
  app.api(route, async () => {
    let buf = vendorCache.get(route);
    if (!buf) {
      buf = await readFile(join(HERE, relPath));
      vendorCache.set(route, buf);
    }
    return new Response(buf, {
      contentType: 'application/javascript; charset=utf-8',
      headers: STATIC_CACHE_HEADERS,
    });
  }, 'GET');
}

// Offline PDF.js CMap endpoint: GET /api/vendor/cmap?file=<name>.bcmap
const cmapCache = new Map();
app.api('/api/vendor/cmap', async (data) => {
  const safeName = sanitizeCmapFileName(data.file);
  if (!safeName) {
    return jsonResponse({ error: '非法的 CMap 文件名' }, 400);
  }

  let buf = cmapCache.get(safeName);
  if (!buf) {
    const fullPath = join(HERE, 'vendor', 'cmaps', safeName);
    if (!existsSync(fullPath)) {
      return jsonResponse({ error: `CMap 文件不存在: ${safeName}` }, 404);
    }
    buf = await readFile(fullPath);
    cmapCache.set(safeName, buf);
  }

  return new Response(buf, {
    contentType: 'application/octet-stream',
    headers: STATIC_CACHE_HEADERS,
  });
}, 'GET');

// ---------------------------------------------------------------------------
// System & Diagnostic API
// ---------------------------------------------------------------------------
app.api('/api/status', () => ({
  name: 'qy-office-viewer',
  version: '1.2.1',
  node: process.version,
  platform: `${platform()} ${release()}`,
  uptimeSeconds: Math.round((Date.now() - STARTED_AT) / 1000),
}), 'GET');

// ---------------------------------------------------------------------------
// Single Active Directory Watcher (fs.watch + 150ms debounce)
// ---------------------------------------------------------------------------
const normWatchKey = (p) => (platform() === 'win32' ? resolve(p).toLowerCase() : resolve(p));

let activeWatch = {
  key: null,
  targetPath: null,
  dirPath: null,
  fileNameLower: null,
  watcher: null,
  debounceTimer: null,
  state: null,
};

function closeActiveWatch() {
  if (activeWatch.debounceTimer) {
    clearTimeout(activeWatch.debounceTimer);
    activeWatch.debounceTimer = null;
  }
  if (activeWatch.watcher) {
    try {
      activeWatch.watcher.close();
    } catch {
      // ignore close errors
    }
    activeWatch.watcher = null;
  }
}

function refreshActiveWatchState() {
  if (!activeWatch.targetPath || !activeWatch.state) return null;
  const st = safeStat(activeWatch.targetPath);
  const prev = activeWatch.state;

  if (!st || !st.isFile) {
    if (prev.exists !== false) {
      activeWatch.state = {
        exists: false,
        version: prev.version + 1,
        mtime: 0,
        size: 0,
      };
    }
  } else if (!prev.exists || st.mtime !== prev.mtime || st.size !== prev.size) {
    activeWatch.state = {
      exists: true,
      version: prev.version + 1,
      mtime: st.mtime,
      size: st.size,
    };
  }
  return activeWatch.state;
}

function ensureWatchingFile(resolvedPath, forceRefreshStat = false) {
  const key = normWatchKey(resolvedPath);
  if (activeWatch.key === key && activeWatch.watcher) {
    if (forceRefreshStat) {
      refreshActiveWatchState();
    }
    return activeWatch.state;
  }

  closeActiveWatch();

  const st = safeStat(resolvedPath);
  const dirPath = dirname(resolvedPath);
  const fileNameLower = basename(resolvedPath).toLowerCase();

  activeWatch = {
    key,
    targetPath: resolvedPath,
    dirPath,
    fileNameLower,
    watcher: null,
    debounceTimer: null,
    state: st && st.isFile
      ? { exists: true, version: 1, mtime: st.mtime, size: st.size }
      : { exists: false, version: 1, mtime: 0, size: 0 },
  };

  try {
    const watcher = watch(dirPath, (_eventType, changedFile) => {
      if (!changedFile || String(changedFile).toLowerCase() === activeWatch.fileNameLower) {
        if (activeWatch.debounceTimer) clearTimeout(activeWatch.debounceTimer);
        activeWatch.debounceTimer = setTimeout(() => {
          activeWatch.debounceTimer = null;
          refreshActiveWatchState();
        }, 150);
      }
    });
    watcher.on('error', (err) => {
      console.warn('[fs.watch] Watcher error:', err.message);
      closeActiveWatch();
    });
    activeWatch.watcher = watcher;
  } catch (err) {
    console.warn('[fs.watch] Could not watch directory:', err.message);
  }

  return activeWatch.state;
}

// ---------------------------------------------------------------------------
// File Content, Meta & Watch State APIs
// ---------------------------------------------------------------------------

// GET /api/file/raw?path=...
app.api('/api/file/raw', async (data) => {
  const rawInput = typeof data.path === 'string' ? data.path.trim().replace(/^["']|["']$/g, '') : '';
  if (!rawInput) {
    return jsonResponse({ error: '缺少文件路径参数' }, 400);
  }

  const actualPath = resolveExactPath(rawInput);
  if (!actualPath) {
    const existingNonFile = existsSync(rawInput) && !safeStat(rawInput)?.isFile;
    if (existingNonFile) {
      return jsonResponse({ error: '目标路径不是普通文件' }, 400);
    }
    return jsonResponse({ error: '文件不存在或路径无效' }, 404);
  }

  if (!isAllowedPreviewExt(actualPath)) {
    return jsonResponse({ error: '不允许预览此类型文件' }, 403);
  }

  const st = safeStat(actualPath);
  if (!st || !st.isFile) {
    return jsonResponse({ error: '目标路径不是有效文件' }, 400);
  }

  const ext = extname(actualPath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  try {
    const buffer = await readFile(actualPath);
    return new Response(buffer, {
      contentType,
      headers: {
        'Content-Disposition': `inline; filename="${encodeURIComponent(basename(actualPath))}"`,
        ...NO_CACHE_HEADERS,
        'Expires': '0',
      },
    });
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}, 'GET');

// GET /api/file/meta?path=...
app.api('/api/file/meta', (data) => {
  const actualPath = resolveExactPath(data.path);
  if (!actualPath) {
    return {
      exists: false,
      error: '文件不存在或路径无效',
      requestedPath: data.path || '',
      candidates: findSimilarCandidates(data.path),
    };
  }

  const ext = extname(actualPath).toLowerCase();
  const allowed = isAllowedPreviewExt(actualPath);
  const st = safeStat(actualPath);

  if (allowed) {
    ensureWatchingFile(actualPath, true);
  }

  return {
    exists: true,
    allowed,
    name: basename(actualPath),
    path: actualPath,
    originalPath: data.path,
    size: st?.size ?? 0,
    mtime: st?.mtime ?? 0,
    ext,
    type: getFileType(ext),
  };
}, 'GET');

// GET /api/file/watch-state?path=...
app.api('/api/file/watch-state', (data) => {
  const rawInput = typeof data.path === 'string' ? data.path.trim().replace(/^["']|["']$/g, '') : '';
  if (!rawInput) {
    return { exists: false, version: 0, mtime: 0, size: 0 };
  }

  const cleanKey = normWatchKey(rawInput);
  if (activeWatch.key === cleanKey && activeWatch.watcher && activeWatch.state) {
    return {
      path: activeWatch.targetPath,
      ...activeWatch.state,
    };
  }

  const actualPath = resolveExactPath(rawInput);
  if (!actualPath) {
    return { exists: false, version: 0, mtime: 0, size: 0 };
  }

  if (!isAllowedPreviewExt(actualPath)) {
    return jsonResponse({ exists: false, error: '不允许预览此类型文件' }, 403);
  }

  const state = ensureWatchingFile(actualPath, false);
  return {
    path: actualPath,
    ...state,
  };
}, 'GET');

// Server-side Spreadsheet Parser (BOM + UTF-8/GB18030 lossless CSV & XLSX)
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
app.api('/api/file/spreadsheet', async (data) => {
  const actualPath = resolveExactPath(data.path);
  if (!actualPath) {
    return jsonResponse({ ok: false, error: '文件不存在或路径无效' }, 404);
  }

  if (!isSpreadsheetExt(actualPath)) {
    return jsonResponse({ ok: false, error: '不支持作为表格解析的文件类型' }, 403);
  }

  try {
    const xlsxLib = getBackendXLSX();
    const buf = await readFile(actualPath);
    const ext = extname(actualPath).toLowerCase();

    const wb = ext === '.csv'
      ? xlsxLib.read(decodeCsvBuffer(buf), { type: 'string' })
      : xlsxLib.read(buf, { type: 'buffer' });

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

    return {
      ok: true,
      sheetNames,
      sheets,
    };
  } catch (err) {
    console.error('[spreadsheet] Backend parse error:', err);
    return jsonResponse({ ok: false, error: err.message }, 500);
  }
}, 'GET');

// ---------------------------------------------------------------------------
// Native System Actions (Safe Spawning without cmd.exe injection)
// ---------------------------------------------------------------------------

function cleanPsStderr(raw) {
  if (!raw) return '';
  let text = String(raw).trim();
  if (text.startsWith('#< CLIXML')) {
    const errMatches = [];
    const re = /<S S="Error">([\s\S]*?)<\/S>/g;
    let m;
    while ((m = re.exec(text)) !== null) {
      errMatches.push(m[1]);
    }
    text = errMatches.length > 0 ? errMatches.join(' ') : text.replace(/<[^>]+>/g, ' ');
    text = text
      .replace(/_x000D__x000A_/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }
  return text.replace(/\s+/g, ' ').trim();
}

function spawnDetachedAndConfirm(command, args, options = {}) {
  return new Promise((resolveResult) => {
    let settled = false;
    const finish = (result) => {
      if (!settled) {
        settled = true;
        resolveResult(result);
      }
    };

    try {
      const child = spawn(command, args, {
        detached: true,
        stdio: 'ignore',
        ...options,
      });

      child.on('error', (err) => {
        console.error(`[spawn] ${command} error:`, err);
        finish({ ok: false, error: err.message });
      });

      child.once('spawn', () => {
        try {
          child.unref();
        } catch {
          // ignore unref errors
        }
        finish({ ok: true });
      });
    } catch (err) {
      finish({ ok: false, error: err.message });
    }
  });
}

function runWinPowerShellAction(buildArgsFn, { timeoutMs = 3000 } = {}) {
  const fallbackExe = resolveWinFallbackPowerShellExe();

  const attempt = (exeOverride, allowFallback) =>
    new Promise((resolveResult) => {
      let settled = false;
      let timer = null;
      let stderrText = '';

      const finish = (result) => {
        if (settled) return;
        settled = true;
        if (timer) {
          clearTimeout(timer);
          timer = null;
        }
        resolveResult(result);
      };

      const { command, args } = buildArgsFn(exeOverride);
      let child;
      try {
        // Must keep detached: false on Windows so console-subsystem pwsh/powershell initializes properly.
        child = spawn(command, args, {
          detached: false,
          stdio: ['ignore', 'ignore', 'pipe'],
          windowsHide: true,
        });
      } catch (err) {
        if (allowFallback && command.toLowerCase() !== fallbackExe.toLowerCase()) {
          finish(attempt(fallbackExe, false));
          return;
        }
        finish({ ok: false, error: err.message });
        return;
      }

      if (child.stderr) {
        child.stderr.on('data', (chunk) => {
          if (!settled && stderrText.length < 2048) {
            stderrText += chunk.toString('utf8').slice(0, 2048 - stderrText.length);
          }
        });
      }

      child.on('error', (err) => {
        console.error(`[ps-action] ${command} error:`, err);
        if (!settled && allowFallback && command.toLowerCase() !== fallbackExe.toLowerCase()) {
          finish(attempt(fallbackExe, false));
          return;
        }
        finish({ ok: false, error: err.message });
      });

      child.on('close', (code) => {
        if (code === 0) {
          finish({ ok: true });
        } else {
          const cleanErr = cleanPsStderr(stderrText) || `PowerShell exited with code ${code}`;
          console.error(`[ps-action] ${command} failed (${code}):`, cleanErr);
          finish({ ok: false, error: cleanErr });
        }
      });

      timer = setTimeout(() => {
        if (!settled) {
          try {
            child.unref();
          } catch {
            // ignore unref errors
          }
          finish({ ok: true });
        }
      }, timeoutMs);
    });

  return attempt(undefined, true);
}

// POST /api/action/open-default { path }
app.api('/api/action/open-default', async (data) => {
  const actualPath = resolveExactPath(data.path);
  if (!actualPath) {
    return { ok: false, error: '文件不存在或路径无效' };
  }

  const absPath = resolve(actualPath);
  const currentPlatform = platform();

  try {
    let result;
    if (currentPlatform === 'win32') {
      result = await runWinPowerShellAction((exe) => buildWinOpenArgs(absPath, exe));
    } else if (currentPlatform === 'darwin') {
      result = await spawnDetachedAndConfirm('open', [absPath]);
    } else {
      result = await spawnDetachedAndConfirm('xdg-open', [absPath]);
    }

    if (!result.ok) {
      return { ok: false, error: result.error || '打开失败' };
    }
    return { ok: true, message: `已在系统默认程序中打开：${basename(absPath)}` };
  } catch (err) {
    console.error('[open-default] Failed:', err);
    return { ok: false, error: err.message };
  }
});

// POST /api/action/reveal { path }
app.api('/api/action/reveal', async (data) => {
  const actualPath = resolveExactPath(data.path);
  if (!actualPath) {
    return { ok: false, error: '文件不存在或路径无效' };
  }

  const absPath = resolve(actualPath);
  const currentPlatform = platform();

  try {
    let result;
    if (currentPlatform === 'win32') {
      const { command, args, verbatim } = buildWinRevealArgs(absPath);
      result = await spawnDetachedAndConfirm(command, args, {
        detached: true,
        stdio: 'ignore',
        windowsVerbatimArguments: !!verbatim,
      });
    } else if (currentPlatform === 'darwin') {
      result = await spawnDetachedAndConfirm('open', ['-R', absPath]);
    } else {
      result = await spawnDetachedAndConfirm('xdg-open', [dirname(absPath)]);
    }

    if (!result.ok) {
      return { ok: false, error: result.error || '定位失败' };
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
      const fallbackExe = resolveWinFallbackPowerShellExe();
      const runClipboard = (exeOverride, allowFallback) => {
        const { command, args } = buildWinClipboardArgs(text, exeOverride);
        execFile(command, args, { windowsHide: true }, (err) => {
          if (err) {
            if (allowFallback && command.toLowerCase() !== fallbackExe.toLowerCase()) {
              runClipboard(fallbackExe, false);
              return;
            }
            console.error('[copy-clipboard] Failed:', err);
            resolveResult({ ok: false, error: err.message });
          } else {
            resolveResult({ ok: true, message: '已成功复制到系统剪贴板' });
          }
        });
      };
      runClipboard(undefined, true);
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

app.run();
