#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_NAME="Billboardy Sync.app"
DIST_DIR="$SCRIPT_DIR/dist"
APP_PATH="$DIST_DIR/$APP_NAME"
TMP_APP_PATH="$DIST_DIR/.Billboardy Sync.tmp.app"
LAUNCHER_PATH="$SCRIPT_DIR/run_sync.sh"

mkdir -p "$DIST_DIR"

if [[ ! -x "$LAUNCHER_PATH" ]]; then
  echo "Missing executable launcher: $LAUNCHER_PATH" >&2
  exit 1
fi

TMP_SCRIPT="$(mktemp "$TMPDIR/billboardy-sync-app.XXXXXX.applescript")"
cleanup() {
  rm -f "$TMP_SCRIPT"
  if [[ -e "$TMP_APP_PATH" ]]; then
    mv "$TMP_APP_PATH" "$TMP_APP_PATH.failed.$(date +%s)"
  fi
}
trap cleanup EXIT

cat > "$TMP_SCRIPT" <<EOF
on run
  set launcherPath to "$(printf '%s' "$LAUNCHER_PATH")"
  try
    set commandText to quoted form of launcherPath
    set runOutput to do shell script commandText
    if runOutput is "" then
      set runOutput to "Synchronizace dokoncena."
    end if
    display dialog runOutput buttons {"OK"} default button "OK"
  on error errorMessage number errorNumber
    display dialog "Synchronizace selhala (" & errorNumber & "):" & return & errorMessage buttons {"OK"} default button "OK" with icon stop
  end try
end run
EOF

if [[ -e "$APP_PATH" ]]; then
  BACKUP_PATH="$DIST_DIR/${APP_NAME}.backup.$(date +%Y%m%d-%H%M%S)"
  mv "$APP_PATH" "$BACKUP_PATH"
  echo "Existing app moved to: $BACKUP_PATH"
fi

osacompile -o "$TMP_APP_PATH" "$TMP_SCRIPT"
mv "$TMP_APP_PATH" "$APP_PATH"

echo "Built app: $APP_PATH"
