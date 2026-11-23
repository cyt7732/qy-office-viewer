#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEST="${HOME}/.gemini/config/plugins/qy-office-viewer"

mkdir -p "${DEST}"
cp -R "${SCRIPT_DIR}/assets" "${SCRIPT_DIR}/sidecars" "${SCRIPT_DIR}/plugin.json" "${SCRIPT_DIR}/README.md" "${SCRIPT_DIR}/CHANGELOG.md" "${DEST}/"
echo "✅ Synced to ${DEST} successfully."
