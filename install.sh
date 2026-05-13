#!/bin/bash
# Install the takimoto-presentation skill to ~/.claude/skills/ for global use.
#
# Usage:
#   bash install.sh           # install
#   bash install.sh --link    # symlink instead of copy (for active development)
#   bash install.sh --uninstall

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_SRC="$SCRIPT_DIR/.claude/skills/takimoto-presentation"
SKILL_DEST="$HOME/.claude/skills/takimoto-presentation"

mode="${1:-install}"

case "$mode" in
  --uninstall)
    if [ -e "$SKILL_DEST" ]; then
      rm -rf "$SKILL_DEST"
      echo "✓ Removed $SKILL_DEST"
    else
      echo "(nothing to remove)"
    fi
    exit 0
    ;;
  --link)
    mkdir -p "$(dirname "$SKILL_DEST")"
    [ -e "$SKILL_DEST" ] && rm -rf "$SKILL_DEST"
    ln -s "$SKILL_SRC" "$SKILL_DEST"
    echo "✓ Symlinked $SKILL_SRC → $SKILL_DEST"
    ;;
  install|"")
    mkdir -p "$(dirname "$SKILL_DEST")"
    [ -e "$SKILL_DEST" ] && rm -rf "$SKILL_DEST"
    cp -r "$SKILL_SRC" "$SKILL_DEST"
    echo "✓ Copied skill to $SKILL_DEST"
    ;;
  *)
    echo "Usage: $0 [install|--link|--uninstall]" >&2
    exit 1
    ;;
esac

echo ""
echo "=== Dependency check ==="

if command -v pdftotext >/dev/null 2>&1; then
  echo "✓ poppler-utils (pdftotext) found"
else
  echo "✗ poppler-utils not found. Install:"
  echo "    macOS:  brew install poppler"
  echo "    Linux:  sudo apt-get install poppler-utils"
fi

if python3 -c "import openpyxl" 2>/dev/null; then
  echo "✓ openpyxl installed"
else
  echo "✗ openpyxl missing. Run: pip3 install openpyxl"
fi

if python3 -c "import pptx" 2>/dev/null; then
  echo "✓ python-pptx installed"
else
  echo "✗ python-pptx missing. Run: pip3 install python-pptx"
fi

echo ""
echo "=== Done ==="
echo "Open a fresh Claude Code session and try one of:"
echo "  /takimoto-presentation"
echo "  「瀧本ゼミ形式で資料作って」"
echo "  「投資推奨のスライド・スプシを作って」"
