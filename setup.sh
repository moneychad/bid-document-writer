#!/bin/bash
# bid-document-writer skill 一键安装脚本
# 用法: bash setup.sh [--full]
#   --full: 安装全部依赖（含LibreOffice和docformat-gui）
#   默认:  仅安装Python最小依赖（覆盖80%功能）

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
FULL_INSTALL=false
PYPI_MIRROR="https://pypi.tuna.tsinghua.edu.cn/simple"

[[ "$1" == "--full" ]] && FULL_INSTALL=true

echo "=== bid-document-writer skill 安装 ==="
echo "模式: $([ "$FULL_INSTALL" = true ] && echo '完整安装' || echo '最小安装')"
echo ""

# 检测 Python
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        PYTHON="$cmd"
        break
    fi
done
if [ -z "$PYTHON" ]; then
    echo "❌ 未找到 Python，请先安装 Python 3.10+"
    exit 1
fi
PY_VER=$($PYTHON -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "✓ Python: $PYTHON ($PY_VER)"

# 检测 pip
PIP="$PYTHON -m pip"
if ! $PIP --version &>/dev/null; then
    echo "❌ pip 不可用，请先安装 pip"
    exit 1
fi
echo "✓ pip: $($PIP --version | head -1)"

# 国内镜像检测
USE_MIRROR=false
if curl -s --connect-timeout 3 "$PYPI_MIRROR" >/dev/null 2>&1; then
    USE_MIRROR=true
    echo "✓ PyPI镜像: $PYPI_MIRROR (可用)"
else
    echo "⚠ PyPI镜像不可用，将使用默认源"
fi
echo ""

# 安装 Python 依赖
echo "--- 安装 Python 依赖 ---"
PIP_ARGS=""
[ "$USE_MIRROR" = true ] && PIP_ARGS="-i $PYPI_MIRROR --trusted-host pypi.tuna.tsinghua.edu.cn"

PACKAGES="python-docx lxml pdfplumber pypdf pymupdf"
echo "安装: $PACKAGES"
$PIP install $PIP_ARGS $PACKAGES 2>&1 | tail -3
echo ""

# 验证安装
echo "--- 验证安装 ---"
ALL_OK=true
for pkg in docx lxml pdfplumber pypdf fitz; do
    if $PYTHON -c "import $pkg" 2>/dev/null || { [ "$pkg" = "fitz" ] && $PYTHON -c "import pymupdf" 2>/dev/null; }; then
        echo "  ✓ $pkg"
    else
        echo "  ❌ $pkg 导入失败"
        ALL_OK=false
    fi
done
echo ""

# 完整安装模式
if [ "$FULL_INSTALL" = true ]; then
    echo "--- 完整安装: LibreOffice ---"
    if command -v soffice &>/dev/null || command -v libreoffice &>/dev/null; then
        echo "  ✓ LibreOffice 已安装"
    elif command -v apt &>/dev/null; then
        echo "  安装 LibreOffice Writer..."
        sudo apt update -qq && sudo apt install -y -qq libreoffice-writer 2>&1 | tail -2
        echo "  ✓ LibreOffice 安装完成"
    elif command -v brew &>/dev/null; then
        echo "  安装 LibreOffice (macOS)..."
        brew install --cask libreoffice 2>&1 | tail -2
        echo "  ✓ LibreOffice 安装完成"
    else
        echo "  ⚠ 无法自动安装LibreOffice，请手动安装"
        echo "    Ubuntu: sudo apt install libreoffice-writer"
        echo "    macOS:  brew install --cask libreoffice"
    fi
    echo ""

    echo "--- 完整安装: docformat-gui ---"
    DOCFORMAT_DIR="/opt/docformat-gui"
    if [ -d "$DOCFORMAT_DIR" ]; then
        echo "  ✓ docformat-gui 已存在于 $DOCFORMAT_DIR"
    else
        # 尝试多个镜像
        CLONED=false
        MIRRORS=(
            "https://ghproxy.com/https://github.com/KaguraNanaga/docformat-gui.git"
            "https://gitclone.com/github.com/KaguraNanaga/docformat-gui.git"
            "https://github.com/KaguraNanaga/docformat-gui.git"
        )
        for url in "${MIRRORS[@]}"; do
            echo "  尝试: $url"
            if git clone --depth 1 "$url" "$DOCFORMAT_DIR" 2>/dev/null; then
                CLONED=true
                echo "  ✓ docformat-gui 克隆成功"
                break
            fi
        done
        if [ "$CLONED" = false ]; then
            echo "  ⚠ docformat-gui 克隆失败（网络问题），跳过"
            echo "    可稍后手动安装: git clone <mirror> $DOCFORMAT_DIR"
        elif [ -f "$DOCFORMAT_DIR/requirements.txt" ]; then
            echo "  安装 docformat-gui 依赖..."
            $PIP install $PIP_ARGS -r "$DOCFORMAT_DIR/requirements.txt" 2>&1 | tail -2
        fi
    fi
    echo ""
fi

# 最终报告
echo "=== 安装完成 ==="
echo ""
echo "已安装的工具:"
echo "  • python-docx   — Word文件读写"
echo "  • lxml          — XML解析（中文字体检查）"
echo "  • pdfplumber    — PDF文本/布局提取（页码检查）"
echo "  • pypdf         — PDF基础操作"
echo "  • pymupdf       — PDF渲染（视觉对比）"

if [ "$FULL_INSTALL" = true ]; then
    echo "  • LibreOffice   — docx→PDF高保真转换"
    echo "  • docformat-gui — 公文格式诊断"
fi

echo ""
if [ "$ALL_OK" = true ]; then
    echo "✅ 全部依赖就绪，skill可正常使用"
else
    echo "⚠ 部分依赖安装失败，请检查上方错误信息"
fi
