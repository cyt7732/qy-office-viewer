# Sync D:\Project\qy-office-viewer to ~/.gemini/config/plugins/qy-office-viewer
$dest = "$env:USERPROFILE\.gemini\config\plugins\qy-office-viewer"
if (-not (Test-Path $dest)) {
    New-Item -ItemType Directory -Path $dest | Out-Null
}
Copy-Item "assets", "sidecars", "plugin.json", "README.md" -Destination $dest -Recurse -Force
Write-Host "✅ Synced to $dest successfully."
