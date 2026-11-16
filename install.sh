#!/usr/bin/env bash
# Install (or update) pages-cli and its Claude skill from GitHub.
# Usage: bash install.sh [install-folder]     (default: ~/pages-cli)
# To install from a fork, set PAGES_CLI_REPO to its clone URL first.
set -euo pipefail

REPO_URL="${PAGES_CLI_REPO:-https://github.com/marcelrgberger/pages-cli.git}"
DEST="${1:-$HOME/pages-cli}"

# Find a Python 3.10 or newer (macOS's built-in python3 is often 3.9).
PYTHON=""
for candidate in python3.14 python3.13 python3.12 python3.11 python3.10 python3; do
  if command -v "$candidate" >/dev/null 2>&1 \
     && "$candidate" -c 'import sys; sys.exit(sys.version_info < (3, 10))' 2>/dev/null; then
    PYTHON="$candidate"
    break
  fi
done
if [ -z "$PYTHON" ]; then
  echo "pages-cli needs Python 3.10 or newer, and none was found." >&2
  echo "Install one with:  brew install python" >&2
  echo "or download it from https://www.python.org/downloads/ , then run this script again." >&2
  exit 1
fi
echo "Using $($PYTHON --version) ($(command -v "$PYTHON"))"

if [ -d "$DEST/.git" ]; then
  git -C "$DEST" pull --ff-only
else
  git clone "$REPO_URL" "$DEST"
fi

cd "$DEST/agent-harness"
rm -rf .venv
"$PYTHON" -m venv .venv
.venv/bin/pip install --quiet -e .

mkdir -p "$HOME/.claude/skills"
if [ -e "$HOME/.claude/skills/pages" ] || [ -L "$HOME/.claude/skills/pages" ]; then
  echo "Skill link already exists: ~/.claude/skills/pages (left as is)"
else
  ln -s "$DEST/skills/pages" "$HOME/.claude/skills/pages"
  echo "Linked skill: ~/.claude/skills/pages"
fi

.venv/bin/pages-cli --help >/dev/null \
  && echo "OK: pages-cli installed in $DEST" \
  || { echo "Install finished but pages-cli does not start" >&2; exit 1; }
echo "Activate it with: source $DEST/agent-harness/.venv/bin/activate"
