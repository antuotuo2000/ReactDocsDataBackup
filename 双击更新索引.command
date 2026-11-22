#!/bin/bash
export PATH="/opt/homebrew/bin:/usr/local/bin:/Library/Frameworks/Python.framework/Versions/Current/bin:$PATH"
cd "$(dirname "$0")" || exit 1
xattr -dr com.apple.quarantine sync_docs_index.py docs 2>/dev/null || true

if command -v python3 >/dev/null 2>&1; then
  PY_BIN="python3"
elif command -v python >/dev/null 2>&1; then
  PY_BIN="python"
else
  echo "❌ 未检测到 Python 3，请先安装 Python 3 (https://www.python.org/)"
  read -n 1 -s -r -p "按任意键关闭窗口..."
  exit 1
fi

"$PY_BIN" sync_docs_index.py --open
echo ""
read -n 1 -s -r -p "按任意键关闭窗口..."
