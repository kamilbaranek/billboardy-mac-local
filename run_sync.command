#!/bin/zsh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
"$SCRIPT_DIR/run_sync.sh" "$@"
STATUS=$?

if [[ -t 1 ]]; then
  printf '\nStiskni Enter pro zavreni...'
  read -r
fi

exit "$STATUS"
