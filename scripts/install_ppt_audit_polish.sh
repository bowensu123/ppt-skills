#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$SCRIPT_DIR/../.opencode/skills/ppt-audit-polish"
TARGET_ROOT="${OPENCODE_SKILL_TARGET_ROOT:-$HOME/.config/opencode/skills}"
TARGET_DIR="$TARGET_ROOT/ppt-audit-polish"

mkdir -p "$TARGET_ROOT"
rm -rf "$TARGET_DIR"
cp -R "$SOURCE_DIR" "$TARGET_DIR"
find "$TARGET_DIR" \( -name '.venv' -o -name '.pytest_cache' -o -name '__pycache__' -o -name '*.egg-info' \) -exec rm -rf {} +

echo "Installed ppt-audit-polish to $TARGET_DIR"
