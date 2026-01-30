#!/usr/bin/env pwsh
# BOM Processor 测试运行脚本
# 快速运行回归测试

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "BOM Processor 回归测试" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# 激活虚拟环境
$venvPath = ".\bom_streamlit\Scripts\Activate.ps1"
if (Test-Path $venvPath) {
    Write-Host "激活虚拟环境..." -ForegroundColor Yellow
    & $venvPath
} else {
    Write-Host "警告: 虚拟环境未找到，使用系统Python" -ForegroundColor Red
}

# 运行测试
Write-Host ""
Write-Host "运行测试..." -ForegroundColor Yellow
Write-Host ""

& python -m pytest tests\test_bom_processor.py -v 

# 显示结果
Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "测试完成" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
