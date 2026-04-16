#!/bin/zsh
set -euo pipefail

PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PID_FILE="$SCRIPT_DIR/.title_mapper_server.pid"
LOG_FILE="$SCRIPT_DIR/.title_mapper_server.log"
ENV_FILE="$SCRIPT_DIR/.env.title_mapper"

if [[ -f "$ENV_FILE" ]]; then
  set -a
  source "$ENV_FILE"
  set +a
fi

# Provide a sensible default for base URL if omitted.
: "${CM_API_BASE_URL:=https://manager.crewcall.hu/api}"

if [[ -z "${CM_API_TOKEN:-}" ]]; then
  osascript -e 'display dialog "Missing CM_API_TOKEN. Put CM_API_TOKEN in .env.title_mapper" buttons {"OK"} default button "OK"'
  exit 1
fi

if [[ -f "$PID_FILE" ]]; then
  EXISTING_PID="$(cat "$PID_FILE" 2>/dev/null || true)"
  if [[ -n "${EXISTING_PID}" ]] && kill -0 "$EXISTING_PID" 2>/dev/null; then
    open "http://127.0.0.1:8000/export"
    osascript -e 'display notification "Server is already running." with title "Title Mapper"'
    exit 0
  fi
fi

nohup env CM_API_BASE_URL="$CM_API_BASE_URL" CM_API_TOKEN="$CM_API_TOKEN" \
  python3 -m uvicorn backend.app:app --host 127.0.0.1 --port 8000 > "$LOG_FILE" 2>&1 &

NEW_PID=$!
echo "$NEW_PID" > "$PID_FILE"

sleep 1
if kill -0 "$NEW_PID" 2>/dev/null; then
  open "http://127.0.0.1:8000/export"
  osascript -e 'display notification "Title Mapper server started." with title "Title Mapper"'
else
  osascript -e 'display dialog "Failed to start server. Check .title_mapper_server.log" buttons {"OK"} default button "OK"'
  exit 1
fi
