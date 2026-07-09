#!/bin/zsh
# 前后端一键重启（macOS）
# 用法: ./restart.sh
set -e

ROOT="$(cd "$(dirname "$0")" && pwd)"
PY="$ROOT/backend/.venv/bin/uvicorn"
NODE="/Users/edy/.workbuddy/binaries/node/versions/22.22.2/bin/npm"

echo "==> 停止旧进程"
for p in $(pgrep -f "uvicorn app.main:app" 2>/dev/null); do kill "$p" 2>/dev/null || true; done
for p in $(pgrep -f "vite" 2>/dev/null); do kill "$p" 2>/dev/null || true; done
sleep 2

echo "==> 启动后端 (8000)"
cd "$ROOT/backend"
nohup .venv/bin/uvicorn app.main:app --host 0.0.0.0 --port 8000 > /tmp/lt-backend.log 2>&1 < /dev/null & disown

echo "==> 启动前端 (5173)"
cd "$ROOT/frontend"
nohup "$NODE" run dev > /tmp/lt-frontend.log 2>&1 < /dev/null & disown

sleep 5
echo "==> 健康检查"
curl -s -m 3 -w "backend /api/health -> %{http_code}\n" http://127.0.0.1:8000/api/health
curl -s -m 3 -o /dev/null -w "frontend /         -> %{http_code}\n" http://127.0.0.1:5173/
curl -s -m 5 -o /dev/null -w "template (proxy)   -> %{http_code}\n" http://127.0.0.1:5173/api/template
echo "==> done"
