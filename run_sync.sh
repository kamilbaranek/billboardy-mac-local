#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"
CONFIG_PATH="${CONFIG_PATH:-$SCRIPT_DIR/config.json}"
PYTHON_BIN="${PYTHON_BIN:-python3}"
STAMP_FILE="$VENV_DIR/.requirements.sha256"
REQUIREMENTS_HASH="$(shasum -a 256 "$SCRIPT_DIR/requirements.txt" | awk '{print $1}')"

if [[ ! -x "$VENV_DIR/bin/python" ]]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

if [[ ! -f "$STAMP_FILE" || "$(cat "$STAMP_FILE")" != "$REQUIREMENTS_HASH" ]]; then
  "$VENV_DIR/bin/python" -m pip install -r "$SCRIPT_DIR/requirements.txt"
  printf '%s\n' "$REQUIREMENTS_HASH" > "$STAMP_FILE"
fi

exec "$VENV_DIR/bin/python" "$SCRIPT_DIR/sync_billboard_occupancy.py" --config "$CONFIG_PATH" "$@"
