#!/bin/bash
# onework Skill 安装脚本
# 自动检测 Python 环境并安装依赖

set -e

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
echo "📦 安装 onework Skill..."
echo "   目录: $SKILL_DIR"

# 检测 Python
if command -v python3 &> /dev/null; then
    PYTHON=python3
elif command -v python &> /dev/null; then
    PYTHON=python
else
    echo "❌ 错误: 未找到 Python，请先安装 Python 3.8+"
    exit 1
fi

echo "   Python: $PYTHON"
$PYTHON --version

# 检查 pip
if ! command -v pip &> /dev/null && ! command -v pip3 &> /dev/null; then
    echo "❌ 错误: 未找到 pip"
    exit 1
fi

PIP=${PIP:-pip3}
if ! command -v $PIP &> /dev/null; then
    PIP=pip
fi

# 安装依赖
echo ""
echo "📥 安装依赖..."
cd "$SKILL_DIR"
$PIP install -r requirements.txt -q

echo ""
echo "✅ onework Skill 安装完成！"
echo ""
echo "可用工具:"
echo "  • read_document    - 读取文档"
echo "  • extract_structure - 提取大纲"
echo "  • convert_format   - 格式转换"
echo "  • create_from_outline - 生成文档"
echo ""
