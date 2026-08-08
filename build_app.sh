#!/bin/bash
# 把 Matrix AI 数据工具箱打包成独立 .app（双击即用，无需装 Python）
#
# 用法:  bash toolbox/build_app.sh
# 产物:  toolbox/dist/MatrixAI数据工具箱.app（内含空白 config.example.json）
#
# 注意: JDK 与 JDBC 驱动不打进包内，启动后在本机页面中设置连接。

set -e
cd "$(dirname "$0")"

PY=/usr/bin/python3
NAME="MatrixAI数据工具箱"

echo "════════════════════════════════════════════"
echo "  打包 $NAME"
echo "════════════════════════════════════════════"

# 预检：必需文件
for f in app.py ai_tools.py matrix_core.py sql_tools.py convert.py excel_diff.py config.example.json templates/index.html; do
  [ -f "$f" ] || { echo "❌ 缺少 $f"; exit 1; }
done
$PY -m PyInstaller --version >/dev/null 2>&1 || {
  echo "❌ 未安装 PyInstaller，先执行: $PY -m pip install --user pyinstaller"; exit 1; }

# 预编译 OCR 二进制随包分发，用户机器不必装 Xcode 命令行工具
mkdir -p vendor
if command -v swiftc >/dev/null 2>&1; then
  echo "→ 预编译 Vision OCR 程序…"
  $PY - <<'PYEOF'
import sys
sys.path.insert(0, ".")
from pathlib import Path
import convert
Path("vendor").mkdir(exist_ok=True)
Path("vendor/ocr.swift").write_text(convert.OCR_SWIFT, encoding="utf-8")
PYEOF
  swiftc -O vendor/ocr.swift -o vendor/ocr && echo "  ✅ vendor/ocr" \
    || echo "  ⚠️ OCR 编译失败，图片识别功能在目标机需要 Xcode CLT"
else
  echo "  ⚠️ 本机无 swiftc，跳过 OCR 预编译"
fi

rm -rf build dist "$NAME.spec"

OCR_DATA=""
[ -f vendor/ocr ] && OCR_DATA="--add-data vendor/ocr:vendor"

$PY -m PyInstaller \
  --name "$NAME" \
  --windowed \
  --noconfirm \
  --clean \
  --add-data "templates:templates" \
  --add-data "config.example.json:." \
  $OCR_DATA \
  --hidden-import openpyxl \
  --hidden-import openpyxl.styles \
  --hidden-import openpyxl.utils \
  --hidden-import fpdf \
  --hidden-import docx \
  --collect-all docx \
  --collect-all fpdf \
  --exclude-module tkinter \
  --exclude-module matplotlib \
  --exclude-module numpy \
  --exclude-module pandas \
  --exclude-module PIL \
  --exclude-module PyQt5 \
  --exclude-module test \
  app.py

APP="dist/$NAME.app"
if [ -d "$APP" ]; then
  SIZE=$(du -sh "$APP" | cut -f1)
  echo ""
  echo "✅ 打包完成"
  echo "   $(pwd)/$APP   ($SIZE)"
  echo ""
  echo "   双击 .app 即可运行，会自动打开浏览器。"
  echo "   运行日志: ~/.matrix_toolbox/运行日志.log"
else
  echo "❌ 打包失败，未生成 $APP"
  exit 1
fi
