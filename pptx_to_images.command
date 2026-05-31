#!/bin/bash
# ============================================================
#  PPTX 转图片器 V1.1.0 - macOS 启动器
#  双击即可运行，兼容 Intel 和 Apple Silicon Mac
# ============================================================
chmod +x "$0" >/dev/null 2>&1 || true
xattr -d com.apple.quarantine "$0" >/dev/null 2>&1 || true

VERSION="V1.1.0"
LAUNCHER_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$LAUNCHER_DIR" || exit 1
PYTHON_SCRIPT="$LAUNCHER_DIR/pptx_to_images.py"

if [[ ! -f "$PYTHON_SCRIPT" ]]; then
  echo "找不到同目录下的 pptx_to_images.py。" >&2
  echo "请把 pptx_to_images.command 和 pptx_to_images.py 放在同一个文件夹。" >&2
  read -r -p "按回车关闭..." _
  exit 1
fi

PYTHON_BIN="${PYTHON_BIN:-}"
if [[ -z "$PYTHON_BIN" ]]; then
  if command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="$(command -v python3)"
  elif [[ -x /usr/bin/python3 ]]; then
    PYTHON_BIN="/usr/bin/python3"
  else
    echo "============================================"
    echo "  需要 Python 3，您的 Mac 尚未安装。"
    echo "  正在尝试通过 Xcode 命令行工具安装..."
    echo "============================================"
    xcode-select --install 2>/dev/null || true
    echo ""
    echo "如果自动安装失败，请手动安装："
    echo "  https://www.python.org/downloads/"
    echo ""
    echo "安装完成后重新双击本文件即可。"
    read -r -p "按回车关闭..." _
    exit 1
  fi
fi

if [[ $# -gt 0 ]]; then
  "$PYTHON_BIN" "$PYTHON_SCRIPT" "$@"
else
  "$PYTHON_BIN" "$PYTHON_SCRIPT"
fi
