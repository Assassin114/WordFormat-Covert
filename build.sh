#!/bin/bash
# ==========================================
# DocFormatter 打包脚本 (Linux / macOS)
# 执行: bash build.sh
# 输出: dist/DocFormatter/ 文件夹
# 发布: 将 dist/DocFormatter/ 打包成 tar.gz
# ==========================================

echo ""
echo "=== DocFormatter 打包工具 ==="
echo ""

if ! python3 -m PyInstaller --version &>/dev/null; then
    echo "[安装 PyInstaller...]"
    pip3 install pyinstaller
fi

echo "[开始打包...]"

# Linux/macOS 的 --add-data 分隔符是 :
python3 -m PyInstaller \
    --noconfirm \
    --name=DocFormatter \
    --windowed \
    --add-data="docformatter/templates/builtin:docformatter/templates/builtin" \
    docformatter/main.py

echo ""
echo "=== 打包完成 ==="
echo "输出目录: dist/DocFormatter/"
echo "Linux 发布: tar -czf DocFormatter-linux.tar.gz dist/DocFormatter/"
echo "macOS 发布: tar -czf DocFormatter-macos.tar.gz dist/DocFormatter/  或打包成 .app"
echo ""
