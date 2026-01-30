#!/bin/bash
# BOM Processor 测试运行脚本 (Linux/Mac)

echo "====================================="
echo "BOM Processor 回归测试"
echo "====================================="
echo ""

# 激活虚拟环境
if [ -f "bom_streamlit/bin/activate" ]; then
    echo "激活虚拟环境..."
    source bom_streamlit/bin/activate
else
    echo "警告: 虚拟环境未找到，使用系统Python"
fi

# 运行测试
echo ""
echo "运行测试..."
echo ""

python -m pytest tests/test_bom_processor.py -v --tb=short

# 显示结果
echo ""
echo "====================================="
echo "测试完成"
echo "====================================="
