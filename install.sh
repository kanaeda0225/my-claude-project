#!/bin/bash
# Install the takimoto skills to ~/.claude/skills/ for global use.
#
# Two skills are installed:
#   - takimoto-slides     : generates .pptx from markdown
#   - takimoto-valuation  : generates .xlsx from markdown + financial PDFs
#
# Usage:
#   bash install.sh           # install (copy)
#   bash install.sh --link    # symlink (for active development)
#   bash install.sh --uninstall

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_SRC="$SCRIPT_DIR/.claude/skills"
SKILLS_DEST="$HOME/.claude/skills"

mode="${1:-install}"

SKILL_NAMES=(takimoto-slides takimoto-valuation)
# Legacy names from earlier development that we now clean up
LEGACY_NAMES=(takimoto-presentation)

# Always clean up legacy names first (migration)
for legacy in "${LEGACY_NAMES[@]}"; do
  legacy_path="$SKILLS_DEST/$legacy"
  if [ -e "$legacy_path" ] || [ -L "$legacy_path" ]; then
    rm -rf "$legacy_path"
    echo "✓ Removed legacy: $legacy_path"
  fi
done

case "$mode" in
  --uninstall)
    for s in "${SKILL_NAMES[@]}"; do
      target="$SKILLS_DEST/$s"
      if [ -e "$target" ]; then
        rm -rf "$target"
        echo "✓ Removed $target"
      fi
    done
    exit 0
    ;;
  --link)
    mkdir -p "$SKILLS_DEST"
    for s in "${SKILL_NAMES[@]}"; do
      src="$SKILLS_SRC/$s"
      dest="$SKILLS_DEST/$s"
      [ -e "$dest" ] && rm -rf "$dest"
      ln -s "$src" "$dest"
      echo "✓ Symlinked $src → $dest"
    done
    ;;
  install|"")
    mkdir -p "$SKILLS_DEST"
    for s in "${SKILL_NAMES[@]}"; do
      src="$SKILLS_SRC/$s"
      dest="$SKILLS_DEST/$s"
      [ -e "$dest" ] && rm -rf "$dest"
      cp -r "$src" "$dest"
      echo "✓ Copied $s to $dest"
    done
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
  echo "✓ openpyxl installed (for .xlsx generation)"
else
  echo "✗ openpyxl missing. Run: pip3 install openpyxl"
fi

if python3 -c "import pptx" 2>/dev/null; then
  echo "✓ python-pptx installed (for .pptx generation)"
else
  echo "✗ python-pptx missing. Run: pip3 install python-pptx"
fi

echo ""
echo "=== Done ==="
echo "Open a fresh Claude Code session and try one of:"
echo "  /takimoto-slides       — to generate slides only"
echo "  /takimoto-valuation    — to generate spreadsheet only"
echo "  「○○のスライド作って」  → takimoto-slides が起動"
echo "  「○○のスプシ作って」    → takimoto-valuation が起動"
