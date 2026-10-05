#!/bin/bash
# Mở ứng dụng web xuất hợp đồng. Bấm đúp vào file này là chạy được.
set -e
cd "$(dirname "$0")/.."

PY=.venv/bin/python
if [ ! -x "$PY" ]; then
  echo "Lần đầu chạy: đang cài đặt môi trường Python…"
  python3 -m venv .venv
  .venv/bin/pip install -q --upgrade pip
  .venv/bin/pip install -q openpyxl python-docx docxtpl reportlab fastapi "uvicorn[standard]" python-multipart
fi

PORT=8777
echo "Đang khởi động máy chủ tại http://127.0.0.1:$PORT"
( sleep 2; open "http://127.0.0.1:$PORT" ) &
exec "$PY" -m uvicorn CONTRACTS.app.server:app --host 127.0.0.1 --port "$PORT"
