#!/bin/bash
# ==============================================================================
# HỆ THỐNG LẬP KẾ HOẠCH KIỂM SOÁT, GIÁM SÁT FECT
# ==============================================================================

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "=================================================================="
echo "  HỆ THỐNG LẬP KẾ HOẠCH KIỂM SOÁT, GIÁM SÁT FECT (QA FECT)"
echo "=================================================================="

# Check venv
if [ ! -d ".venv" ]; then
    echo "[1/3] Đang khởi tạo môi trường Python ảo (.venv)..."
    python3 -m venv .venv
    echo "[2/3] Đang cài đặt thư viện cần thiết..."
    .venv/bin/pip install -r requirements.txt
else
    echo "[1/2] Đã tìm thấy môi trường ảo .venv."
fi

# Determine python runner
PYTHON_BIN=".venv/bin/python3"
if [ ! -f "$PYTHON_BIN" ]; then
    PYTHON_BIN="python3"
fi

echo "[2/2] Đang khởi động máy chủ Web tại http://localhost:8000 ..."

# Open browser in background after 1.5s
(sleep 1.5 && open "http://localhost:8000" 2>/dev/null || true) &

# Run server
PYTHONPATH=.venv/lib/python3.13/site-packages python3 app.py
