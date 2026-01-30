# BOM 层级处理器 - Streamlit Web 应用

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.31.0-red.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

基于 Streamlit 的 BOM (Bill of Materials) 层级处理和比较 Web 应用程序。支持从 Excel 文件提取 Level 2 项目、创建压缩表、生成 BOM 汇总，以及比较两个 BOM 文件。

## ✨ 功能特性

### BOM 层级处理

### BOM 文件比较


## 🏗️ 项目结构

```
BOM_Summary/
├── bom_streamlit/              # 虚拟环境
├── streamlit_app/              # Web 应用
│   ├── bom_app.py              # 主应用程序
│   ├── config.py               # 配置文件
│   ├── pages/
│   │   └── 1_BOM_Comparison.py # BOM 比较页面
│   └── utils/
│       ├── __init__.py
│       └── cleanup.py          # 清理工具
├── tests/                      # 测试套件 ⭐ 新增
│   ├── test_bom_processor.py   # 回归测试
│   ├── test_data/
│   │   ├── baseline/           # 基准数据（不提交到Git）
│   │   └── temp_outputs/       # 临时测试输出
│   └── README.md               # 测试文档
├── temp_processing/            # 临时文件（自动清理）
├── data/                       # 数据文件
├── bom_processor.py            # BOM 层级处理核心模块（CLI/Web）
├── bom_comparison.py           # BOM 文件比较核心模块（CLI/Web）
├── requirements.txt            # Python 依赖
├── pytest.ini                  # Pytest 配置 ⭐ 新增
├── .gitignore                  # Git 忽略文件
├── run_app.bat                 # Windows 启动脚本
├── run_app.sh                  # Linux/Mac 启动脚本
├── run_tests.ps1               # Windows 测试脚本 ⭐ 新增
├── run_tests.sh                # Linux/Mac 测试脚本 ⭐ 新增
├── DEPLOYMENT.md               # 详细部署文档
└── README.md                   # 本文件
```

## 🚀 快速开始

### 前置要求

- Python 3.8 或更高版本
- pip（Python 包管理器）

### 安装步骤

1. **克隆或下载项目**
   ```bash
   cd C:\Users\zongyue.liu\Desktop\AdvSimulation\BOM_Summary
   ```

2. **创建虚拟环境**（如果还没有）
   ```bash
   python -m venv bom_streamlit
   ```

3. **激活虚拟环境**
   
   **Windows (PowerShell):**
   ```powershell
   .\bom_streamlit\Scripts\Activate.ps1
   ```
   
   **Windows (CMD):**
   ```cmd
   bom_streamlit\Scripts\activate.bat
   ```
   
   **Linux/Mac:**
   ```bash
   source bom_streamlit/bin/activate
   ```

4. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

### 运行应用

**方式1: 使用启动脚本（推荐）**

**Windows:**
```cmd
run_app.bat
```

**Linux/Mac:**
```bash
chmod +x run_app.sh
./run_app.sh
```

**方式2: 手动启动**
```bash
# 确保虚拟环境已激活
streamlit run streamlit_app/bom_app.py --server.port=8501 --server.address=0.0.0.0
```

应用将在以下地址启动：
- 本地访问: http://localhost:8501
- 网络访问: http://\<your-ip\>:8501

## ⚙️ 配置

编辑 `streamlit_app/config.py` 可以自定义设置：

```python
# 最大上传文件大小 (MB)
MAX_UPLOAD_SIZE_MB = 100

# 临时文件自动清理时间 (小时)
CLEANUP_HOURS = 24

# 必需列
REQUIRED_COLUMNS = [
    'BOM Line',
    'CAD OEM Part Number',
    'CAD OEM Rev',
    'Quantity'
]
```

## 🖥️ 虚拟机部署

详细的虚拟机部署说明请参考 [DEPLOYMENT.md](DEPLOYMENT.md)，包括：

- Windows/Linux 虚拟机配置
- 防火墙设置
- 系统服务配置
- 自动启动设置
- 故障排查指南

### 快速部署摘要

1. **配置防火墙**（允许端口 8501）
   ```powershell
   # Windows PowerShell (管理员)
   New-NetFirewallRule -DisplayName "BOM Streamlit App" -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow
   ```

2. **启动应用**
   ```cmd
   run_app.bat
   ```

3. **从其他机器访问**
   ```
   http://<VM_IP>:8501
   ```

## 🔧 命令行模式

### BOM 层级处理

原有的命令行功能保持不变，可以直接运行：

```bash
python bom_processor.py
```

这将使用硬编码的输入文件路径处理 BOM。

### BOM 文件比较

`bom_comparison.py` 可以作为独立的命令行工具使用，无需启动 Web 应用：

**基本用法:**
```bash
python bom_comparison.py --file1 data/BOM1.xlsx --file2 data/BOM2.xlsx
```

**指定输出文件:**
```bash
python bom_comparison.py --file1 BOM1.xlsx --file2 BOM2.xlsx --output comparison_result.xlsx
```

**同时指定输出和日志文件:**
```bash
python bom_comparison.py --file1 BOM1.xlsx --file2 BOM2.xlsx --output result.xlsx --log comparison.log
```

**查看帮助:**
```bash
python bom_comparison.py --help
```

**输出说明:**
- 自动生成带时间戳的比较报告（默认：`comparison_result_YYYYMMDD_HHMMSS.xlsx`）
- 自动生成日志文件（默认：`comparison_log_YYYYMMDD_HHMMSS.txt`）
- 比较报告包含 **5 个工作表**：
  1. **NoChange**: 完全相同的零件（Part Number + Rev + 所有属性都一致）
  2. **Part_Rev_NoChange**: Part Number + Rev 相同，但 Material Spec / CAD Oem Name / Thickness 有差异
  3. **RevChange**: Part Number 相同，Rev 不同
  4. **PartChange**: 仅在一个文件中出现的零件
  5. **Summary**: 所有分类汇总（按类别分段显示）

**比较逻辑:**
- **分类规则**：
  - NoChange: Part Number + Rev + 所有属性完全一致
  - Part_Rev_NoChange: Part Number + Rev 一致，其他属性有差异
  - RevChange: Part Number 一致，Rev 不同
  - PartChange: Part Number 在另一文件中找不到
  
- **输出格式**：
  - 并排显示：File1 所有列 | File2 所有列
  - 差异单元格用黄色背景高亮
  - Summary 工作表包含所有类别，类别之间用空行分隔

**比较的列:**
- `CAD OEM Part Number`
- `CAD OEM Rev`
- `Material Spec`
- `CAD Oem Name`
- `Thickness`

## 📊 处理说明

### BOM 层级处理

应用执行以下步骤：

1. **加载工作簿**: 读取 Excel 文件，提取层级结构
2. **分析层级**: 识别 Level 1 和 Level 2 项目
3. **提取分支**: 将每个 Level 2 项目及其子项提取到独立的工作表
4. **创建压缩表**: 对每个分支进行去重和数量汇总
5. **生成汇总**: 创建包含所有零件的统一 BOM 汇总表

**提取工作簿包含:**
- 每个 Level 2 项目的原始工作表（例如 `BA`, `BD`, `CA` 等）
- 对应的压缩工作表（例如 `BA_Compress`, `BD_Compress` 等）

### BOM 文件比较

应用执行以下步骤：

1. **加载两个文件**: 读取并验证必需的列
2. **智能分类**: 根据 Part Number 和 Rev 进行四类分类
3. **生成报告**: 创建 5 个工作表的详细比较报告
4. **高亮差异**: 自动标记所有差异单元格（黄色）

**比较报告特点:**
- 并排对比格式，一目了然
- 智能分类，快速定位问题
- Summary 汇总表，完整视图
- 详细日志，可追溯性强

## 📝 常见问题

**Q: 为什么无法从其他机器访问？**

A: 检查以下几点：
1. 防火墙是否允许端口 8501
2. 应用是否使用 `--server.address=0.0.0.0` 启动
3. 虚拟机网络配置是否正确

**Q: 文件上传失败怎么办？**

A: 可能原因：
- 文件大小超过 100MB 限制
- 文件格式不正确（需要 .xlsm 或 .xlsx）
- 检查 `temp_processing/` 目录权限

**Q: 如何增加上传大小限制？**

A: 编辑 `streamlit_app/config.py`，修改 `MAX_UPLOAD_SIZE_MB` 的值

**Q: BOM 比较的分类逻辑是什么？**

A: 比较分为 4 个类别：
1. **NoChange**: 所有内容完全相同
2. **Part_Rev_NoChange**: Part+Rev 相同，但 Material Spec / CAD Oem Name / Thickness 有变化
3. **RevChange**: Part Number 相同，Rev 不同（可能是版本更新）
4. **PartChange**: Part Number 完全不匹配（新增或删除的零件）

## 🧪 测试

本项目包含完整的回归测试框架，确保代码修改不会破坏现有功能。

### 快速运行测试

**Windows:**
```powershell
.\run_tests.ps1
```

**Linux/Mac:**
```bash
./run_tests.sh
```

### 手动运行测试

```bash
# 激活虚拟环境
.\bom_streamlit\Scripts\Activate.ps1  # Windows
# 或
source bom_streamlit/bin/activate     # Linux/Mac

# 运行所有测试
pytest tests/test_bom_processor.py -v

# 运行特定测试
pytest tests/test_bom_processor.py::TestBOMProcessor::test_output_file_cell_by_cell_comparison -v -s
```

### 测试特性

- ✅ **逐单元格比对**: 验证输出文件的每个单元格数据
- ✅ **浮点数容差**: 自动处理浮点精度差异（默认 1e-6）
- ✅ **统计验证**: 确保 Level 1/2 数量、工作表结构保持一致
- ✅ **性能测试**: 监控处理时间，防止性能退化
- ✅ **详细报告**: 自动生成差异报告（JSON 格式）

### 首次设置测试

1. 复制输入文件到基准目录：
   ```powershell
   Copy-Item "data\BOM-COS1000334701-BA.xlsm" "tests\test_data\baseline\"
   ```

2. 运行测试生成基准文件（首次会自动生成）：
   ```bash
   pytest tests/test_bom_processor.py -v
   ```

3. 查看详细测试文档：
   ```
   tests/README.md
   ```

### 修改代码后的工作流

1. **修改前运行测试** - 确保起点正确
2. **修改代码** - 实现新功能或修复bug
3. **修改后运行测试** - 验证没有引入回归问题
4. **如果测试失败** - 查看差异报告，决定是修复代码还是更新基准

📖 **完整测试文档**: 参见 `tests/README.md`

## 📄 许可证

MIT License

---

**作者**: Zongyue Liu  
**创建日期**: 2026-01-27  
**最后更新**: 2026-01-29  
**版本**: 1.1

**更新日志 (v1.1)**:
- ✨ 新增完整的回归测试框架
- ✨ 支持逐单元格比对和浮点数容差验证
- ✨ 添加测试运行脚本 (`run_tests.ps1` / `run_tests.sh`)
- 📚 更新文档，增加测试部分说明
