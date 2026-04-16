#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PID_FILE="$SCRIPT_DIR/.title_mapper_server.pid"

stopped=0
if [[ -f "$PID_FILE" ]]; then
  PID="$(cat "$PID_FILE" 2>/dev/null || true)"
  if [[ -n "${PID}" ]] && kill -0 "$PID" 2>/dev/null; then
    kill "$PID" 2>/dev/null || true
    sleep 0.5
    if kill -0 "$PID" 2>/dev/null; then
      kill -9 "$PID" 2>/dev/null || true
    fi
    stopped=1
  fi
  rm -f "$PID_FILE"
fi

# Fallback: stop any matching uvicorn process
if pgrep -f "uvicorn backend.app:app --host 127.0.0.1 --port 8000" >/dev/null 2>&1; then
  pkill -f "uvicorn backend.app:app --host 127.0.0.1 --port 8000" || true
  stopped=1
fi

if [[ "$stopped" -eq 1 ]]; then
  osascript -e 'display notification "Title Mapper server stopped." with title "Title Mapper"'
else
  osascript -e 'display notification "No running Title Mapper server found." with title "Title Mapper"'
fi
