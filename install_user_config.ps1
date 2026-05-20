# install_user_config.ps1 — One-shot user-level Claude config install
# Usage (Windows PowerShell): .\install_user_config.ps1
# Run from the repo root. Copies .claude-user/ contents to ~/.claude/

$ErrorActionPreference = "Stop"

$src = Join-Path $PSScriptRoot ".claude-user"
$dst = Join-Path $env:USERPROFILE ".claude"

if (-not (Test-Path $src)) {
    Write-Error "Source directory not found: $src"
    Write-Error "Run this from the repo root (where .claude-user/ lives)."
    exit 1
}

Write-Host "Installing user-level Claude config..." -ForegroundColor Cyan
Write-Host "  Source: $src"
Write-Host "  Destination: $dst"
Write-Host ""

# Make destination directories
New-Item -ItemType Directory -Force -Path $dst                         | Out-Null
New-Item -ItemType Directory -Force -Path "$dst\skills"                | Out-Null
New-Item -ItemType Directory -Force -Path "$dst\templates\ausenco"     | Out-Null
New-Item -ItemType Directory -Force -Path "$dst\scripts"               | Out-Null

# Copy files
Copy-Item -Force "$src\CLAUDE.md"                                   "$dst\CLAUDE.md"
Copy-Item -Force "$src\skills\*"                                    "$dst\skills\"
Copy-Item -Force "$src\templates\ausenco\*"                         "$dst\templates\ausenco\"
Copy-Item -Force "$src\scripts\*.py"                                "$dst\scripts\"

Write-Host "Installed files:" -ForegroundColor Green
Get-ChildItem -Recurse $dst | ForEach-Object {
    if ($_.PSIsContainer) { return }
    $rel = $_.FullName.Substring($dst.Length + 1)
    Write-Host "  $rel  ($([Math]::Round($_.Length / 1024, 1)) KB)"
}

Write-Host ""
Write-Host "Done. Open any new Claude Code session in any workspace to verify." -ForegroundColor Green
Write-Host "Test prompt: 'What are my active engagements?' — Claude should know without you loading anything."
