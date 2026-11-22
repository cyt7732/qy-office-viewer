// QY Office Viewer - Pure Utility & Security Helpers
// Shared between main.mjs and unit verification scripts.

import { existsSync, readdirSync, statSync } from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';

export const ALLOWED_PREVIEW_EXTS = new Set([
  '.docx',
  '.doc',
  '.xlsx',
  '.xls',
  '.csv',
  '.pdf',
  '.pptx',
  '.ppt',
  '.txt',
  '.md',
  '.json',
]);

export const SPREADSHEET_EXTS = new Set(['.xlsx', '.xls', '.csv']);

export const MIME_TYPES = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pdf': 'application/pdf',
  '.doc': 'application/msword',
  '.xls': 'application/vnd.ms-excel',
  '.csv': 'text/csv; charset=utf-8',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.txt': 'text/plain; charset=utf-8',
  '.md': 'text/markdown; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
};

export function getFileType(ext) {
  switch ((ext || '').toLowerCase()) {
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
    case '.txt':
    case '.md':
    case '.json':
      return 'text';
    default:
      return 'other';
  }
}

export function isAllowedPreviewExt(filePath) {
  if (!filePath || typeof filePath !== 'string') return false;
  return ALLOWED_PREVIEW_EXTS.has(extname(filePath).toLowerCase());
}

export function isSpreadsheetExt(filePath) {
  if (!filePath || typeof filePath !== 'string') return false;
  return SPREADSHEET_EXTS.has(extname(filePath).toLowerCase());
}

export function safeStat(filePath) {
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

/**
 * Resolves only an exact file match (allowing case-insensitive basename match
 * on case-sensitive filesystems). Never substitutes a different/similar file.
 */
export function resolveExactPath(inputPath) {
  if (!inputPath || typeof inputPath !== 'string') return null;
  const targetPath = inputPath.trim().replace(/^["']|["']$/g, '');
  if (!targetPath) return null;

  if (existsSync(targetPath)) {
    const st = safeStat(targetPath);
    if (st?.isFile) {
      return resolve(targetPath);
    }
    return null;
  }

  try {
    const dir = dirname(targetPath);
    if (!existsSync(dir)) return null;
    const targetLower = basename(targetPath).toLowerCase();
    const entries = readdirSync(dir, { withFileTypes: true });
    for (const ent of entries) {
      if (ent.isFile() && ent.name.toLowerCase() === targetLower) {
        return resolve(join(dir, ent.name));
      }
    }
  } catch {
    // ignore directory read error
  }

  return null;
}

/**
 * Finds similar previewable files in the same directory for user confirmation
 * when the exact target path does not exist.
 */
export function findSimilarCandidates(inputPath, maxResults = 5) {
  if (!inputPath || typeof inputPath !== 'string') return [];
  const targetPath = inputPath.trim().replace(/^["']|["']$/g, '');
  if (!targetPath) return [];

  try {
    const dir = dirname(targetPath);
    if (!existsSync(dir)) return [];

    const targetBase = basename(targetPath);
    const targetExt = extname(targetBase).toLowerCase();
    const targetStem = basename(targetBase, extname(targetBase)).toLowerCase();
    const cleanTarget = targetStem.replace(/[\s_\-\(\)（）【】\[\]]/g, '');
    if (!cleanTarget) return [];

    const targetChars = Array.from(cleanTarget);
    const entries = readdirSync(dir, { withFileTypes: true });
    const candidates = [];

    for (const ent of entries) {
      if (!ent.isFile()) continue;
      if (!isAllowedPreviewExt(ent.name)) continue;

      const fExt = extname(ent.name).toLowerCase();
      const fStem = basename(ent.name, extname(ent.name)).toLowerCase();
      const cleanCandidate = fStem.replace(/[\s_\-\(\)（）【】\[\]]/g, '');
      if (!cleanCandidate) continue;

      let matches = 0;
      const candidateChars = Array.from(cleanCandidate);
      for (const ch of targetChars) {
        const idx = candidateChars.indexOf(ch);
        if (idx !== -1) {
          matches++;
          candidateChars.splice(idx, 1);
        }
      }

      const similarity = (2 * matches) / (cleanTarget.length + cleanCandidate.length);
      if (similarity < 0.45) continue;

      const fullPath = resolve(join(dir, ent.name));
      const st = safeStat(fullPath);
      if (!st) continue;

      candidates.push({
        name: ent.name,
        path: fullPath,
        size: st.size,
        mtime: st.mtime,
        ext: fExt,
        type: getFileType(fExt),
        similarity: Number(similarity.toFixed(4)),
        _sameExt: fExt === targetExt ? 1 : 0,
      });
    }

    candidates.sort((a, b) => (
      (b.similarity - a.similarity) ||
      (b._sameExt - a._sameExt) ||
      (b.mtime - a.mtime)
    ));

    return candidates.slice(0, maxResults).map(({ _sameExt, ...rest }) => rest);
  } catch {
    return [];
  }
}

/**
 * Decodes CSV raw bytes with BOM inspection, strict UTF-8 validation,
 * and automatic GB18030 fallback for legacy Windows/Excel Chinese CSVs.
 */
export function decodeCsvBuffer(buf) {
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  if (bytes.length === 0) return '';

  // UTF-8 BOM
  if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }

  // UTF-16LE BOM
  if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  }

  // UTF-16BE BOM (swap pairs into LE copy)
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

  // Strict UTF-8 check, fallback to GB18030
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('gb18030').decode(bytes);
  }
}

/**
 * Validates and sanitizes CMap filenames to prevent directory traversal.
 */
export function sanitizeCmapFileName(name) {
  if (typeof name !== 'string' || !name) return null;
  if (name.includes('..') || name.includes('/') || name.includes('\\')) return null;
  if (!/^[A-Za-z0-9._-]+\.bcmap$/.test(name)) return null;
  return name;
}

// ---------------------------------------------------------------------------
// Safe Windows PowerShell Command Builders (-EncodedCommand + single quotes)
// ---------------------------------------------------------------------------
let cachedPsExe = null;

export function resolveWinPowerShellExe() {
  if (cachedPsExe) return cachedPsExe;

  const pathEntries = (process.env.PATH || '').split(';').filter(Boolean);
  for (const dir of pathEntries) {
    const candidate = join(dir, 'pwsh.exe');
    if (existsSync(candidate)) {
      cachedPsExe = candidate;
      return cachedPsExe;
    }
  }

  const programFilesPwsh = 'C:\\Program Files\\PowerShell\\7\\pwsh.exe';
  if (existsSync(programFilesPwsh)) {
    cachedPsExe = programFilesPwsh;
    return cachedPsExe;
  }

  cachedPsExe = 'powershell.exe';
  return cachedPsExe;
}

export function escapePsSingleQuoted(str) {
  return String(str).replace(/'/g, "''");
}

export function encodePsCommand(script) {
  return Buffer.from(String(script), 'utf16le').toString('base64');
}

function buildWinPsInvocation(script) {
  return {
    command: resolveWinPowerShellExe(),
    args: [
      '-NoProfile',
      '-NonInteractive',
      '-WindowStyle',
      'Hidden',
      '-EncodedCommand',
      encodePsCommand(script),
    ],
    script,
  };
}

export function buildWinOpenArgs(absPath) {
  const escaped = escapePsSingleQuoted(absPath);
  return buildWinPsInvocation(`Start-Process -FilePath '${escaped}'`);
}

export function buildWinRevealArgs(absPath) {
  const cleanPath = String(absPath).replace(/"/g, '');
  const escaped = escapePsSingleQuoted(cleanPath);
  return buildWinPsInvocation(
    `Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + '${escaped}' + '"')`
  );
}

export function buildWinClipboardArgs(text) {
  const escaped = escapePsSingleQuoted(text);
  return buildWinPsInvocation(`Set-Clipboard -Value '${escaped}'`);
}
