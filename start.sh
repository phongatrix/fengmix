#!/usr/bin/env bash
# Feng Mix - Start Script
# Usage: bash start.sh [port]

PORT=${1:-8000}
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND="$SCRIPT_DIR/backend"

echo "🎓 Feng Mix - Trộn đề trắc nghiệm Online"
echo "=========================================="

# Check Python 3
if ! command -v python3 &>/dev/null; then
  echo "❌ Cần Python 3.9+. Vui lòng cài đặt Python."
  exit 1
fi

# Install deps if needed
echo "📦 Kiểm tra dependencies..."
cd "$BACKEND"
pip install -r requirements.txt --break-system-packages -q 2>/dev/null || \
  pip install -r requirements.txt -q 2>/dev/null || true

echo ""
echo "🚀 Khởi động server tại http://localhost:$PORT"
echo "   Mở trình duyệt và truy cập: http://localhost:$PORT"
echo "   Nhấn Ctrl+C để dừng server"
echo ""

python3 -m uvicorn main:app --host 0.0.0.0 --port "$PORT" --reload
