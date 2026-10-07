#!/usr/bin/env bash
set -euo pipefail
if command -v realpath >/dev/null 2>&1; then
  SCRIPT_DIR="$(dirname "$(realpath "$0")")"
else
  SCRIPT_DIR="$(dirname "$(readlink "$0" || echo "$0")")"
fi
TMUX_ROOT="$SCRIPT_DIR/tmux-root"
export LD_LIBRARY_PATH="$TMUX_ROOT/lib${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"
export TERMINFO_DIRS="$TMUX_ROOT/share/terminfo${TERMINFO_DIRS:+:$TERMINFO_DIRS}"
exec "$TMUX_ROOT/bin/tmux" "$@"
