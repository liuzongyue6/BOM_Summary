# BOM 层级处理器 - Streamlit Web 应用

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.31.0-red.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

基于 Streamlit 的 BOM (Bill of Materials) 层级处理 Web 应用程序。支持从 Excel 文件提取 Level 2 项目、创建压缩表和生成 BOM 汇总。

## ✨ 功能特性

- 📁 **Web 文件上传**: 通过浏览器上传 BOM Excel 文件 (.xlsm, .xlsx)
- 🔄 **实时处理日志**: 在界面上实时显示处理进度和详细日志
- 📊 **三文件输出**:
  - 提取工作簿 - 包含所有 Level 2 分支和压缩表
  - BOM 汇总 - 合并的零件汇总表
  - 处理日志 - 详细的处理过程记录
- ✅ **输入验证**: 自动检查必需列，给出警告提示
- 🧹 **自动清理**: 24小时自动清理临时文件
- 📈 **统计摘要**: 显示处理结果统计信息

## 🏗️ 项目结构

```
BOM_Summary/
├── bom_streamlit/              # 虚拟环境
├── streamlit_app/              # Web 应用
│   ├── bom_app.py              # 主应用程序
│   ├── config.py               # 配置文件
│   └── utils/
│       ├── __init__.py
│       └── cleanup.py          # 清理工具
├── temp_processing/            # 临时文件（自动清理）
├── data/                       # 数据文件
├── bom_processor.py            # 核心处理逻辑（命令行也可用）
├── requirements.txt            # Python 依赖
├── .gitignore                  # Git 忽略文件
├── run_app.bat                 # Windows 启动脚本
├── run_app.sh                  # Linux/Mac 启动脚本
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

## 📖 使用指南

### 1. 上传文件
- 点击 "浏览文件" 或拖放 Excel 文件到上传区
- 支持 `.xlsm` 和 `.xlsx` 格式
- 最大文件大小：100MB

### 2. 验证文件
应用会自动验证文件结构，检查必需的列：
- BOM Line
- CAD OEM Part Number
- CAD OEM Rev
- Quantity

### 3. 开始处理
点击 "🚀 开始处理" 按钮，实时日志将显示处理进度

### 4. 下载结果
处理完成后，可以下载三个文件：
- **📊 提取工作簿**: 包含所有提取和压缩的工作表
- **📈 BOM 汇总**: 统一的零件汇总表
- **📄 处理日志**: 详细的处理日志文件

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

原有的命令行功能保持不变，可以直接运行：

```bash
python bom_processor.py
```

这将使用硬编码的输入文件路径处理 BOM。

## 📊 处理说明

应用执行以下步骤：

1. **加载工作簿**: 读取 Excel 文件，提取层级结构
2. **分析层级**: 识别 Level 1 和 Level 2 项目
3. **提取分支**: 将每个 Level 2 项目及其子项提取到独立的工作表
4. **创建压缩表**: 对每个分支进行去重和数量汇总
5. **生成汇总**: 创建包含所有零件的统一 BOM 汇总表

### 输出说明

**提取工作簿包含:**
- 每个 Level 2 项目的原始工作表（例如 `BA`, `BD`, `CA` 等）
- 对应的压缩工作表（例如 `BA_Compress`, `BD_Compress` 等）

**BOM 汇总包含:**
- 单个 `BOM_Summary` 工作表
- 所有唯一零件及其在各分支中的数量

## 🛡️ 安全建议

- 仅在受信任的网络环境中运行
- 考虑添加身份验证（使用 streamlit-authenticator）
- 使用 HTTPS（通过 Nginx 反向代理）
- 配置防火墙 IP 白名单

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

**Q: 处理过程中出错怎么办？**

A: 
1. 查看实时日志中的错误信息
2. 下载错误日志进行分析
3. 确认 Excel 文件包含所有必需列
4. 检查文件是否损坏

## 🔄 更新日志

### v1.0 (2026-01-27)
- ✨ 初始 Streamlit Web 版本发布
- 🔄 支持文件上传和实时处理
- 📊 三文件下载功能
- 🧹 自动临时文件清理
- ✅ 输入文件验证
- 📝 实时日志显示

## 👥 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT License

## 🙏 致谢

- [Streamlit](https://streamlit.io/) - Web 应用框架
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Excel 文件处理

---

**作者**: GitHub Copilot  
**创建日期**: 2026-01-27  
**版本**: 1.0
