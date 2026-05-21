#!/usr/bin/env bash
# install_user_config.sh — One-shot user-level Claude config install (Mac / Linux)
# Usage: ./install_user_config.sh
# Run from the repo root. Copies .claude-user/ contents to ~/.claude/

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC="$REPO_ROOT/.claude-user"
DST="$HOME/.claude"

if [[ ! -d "$SRC" ]]; then
    echo "ERROR: source directory not found: $SRC"
    echo "Run this script from the repo root (where .claude-user/ lives)."
    exit 1
fi

echo "Installing user-level Claude config..."
echo "  Source: $SRC"
echo "  Destination: $DST"
echo

mkdir -p "$DST/skills" "$DST/templates/ausenco" "$DST/scripts" \
         "$DST/projects/goldboro" "$DST/projects/green-bay" \
         "$DST/projects/goderich" "$DST/projects/kemess"

cp -f "$SRC/CLAUDE.md"                                "$DST/CLAUDE.md"
cp -f "$SRC/skills/"*                                 "$DST/skills/"
cp -f "$SRC/templates/ausenco/"*                      "$DST/templates/ausenco/"
cp -f "$SRC/scripts/"*.py                             "$DST/scripts/"
cp -f "$SRC/projects/goldboro/"*                      "$DST/projects/goldboro/"
cp -f "$SRC/projects/green-bay/"*                     "$DST/projects/green-bay/"
cp -f "$SRC/projects/goderich/"*                      "$DST/projects/goderich/"
cp -f "$SRC/projects/kemess/"*                        "$DST/projects/kemess/"
chmod +x "$DST/scripts/fd_audit.py"

echo "Installed files:"
find "$DST" -type f -printf "  %P  (%s bytes)\n" | sort

echo
echo "Done. Open any new Claude Code session in any workspace to verify."
echo "Test prompt: 'What are my active engagements?' — Claude should know"
echo "without you loading anything."
