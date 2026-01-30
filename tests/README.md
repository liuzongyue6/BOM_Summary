# BOM Processor 测试指南

## 快速开始

### 首次运行（生成基准）

```powershell
# 1. 激活虚拟环境
.\bom_streamlit\Scripts\Activate.ps1

# 2. 复制输入文件
Copy-Item "data\BOM-COS1000334701-BA.xlsm" "tests\test_data\baseline\"

# 3. 运行测试（自动生成基准）
.\run_tests.ps1
```

### 日常测试

```powershell
# 修改代码后运行
.\run_tests.ps1

# 或使用 pytest 直接运行
pytest tests/test_bom_processor.py -v -s
```

## 测试内容

| 测试项 | 验证内容 |
|--------|---------|
| `test_baseline_files_exist` | 基准文件存在（首次运行自动生成） |
| `test_processor_execution` | 处理器成功运行 |
| `test_statistics_validation` | 统计数据一致性（Level 1/2数量、工作表数） |
| `test_output_file_cell_by_cell_comparison` | **核心**：输出文件逐单元格比对 |
| `test_summary_file_cell_by_cell_comparison` | **核心**：BOM汇总文件逐单元格比对 |
| `test_processing_performance` | 性能测试（3次运行平均，阈值60秒） |

## 配置说明

### pytest.ini 配置文件
- 设置测试发现规则（测试文件、类、函数命名）
- 配置日志输出格式
- 定义测试标记（slow、regression、unit等）

### 测试参数调整

编辑 `tests/test_bom_processor.py`：

```python
class BOMTestConfig:
    FLOAT_TOLERANCE = 1e-6        # 浮点数比较容差
    MAX_PROCESSING_TIME = 60      # 性能阈值（秒）
```

**浮点数容差说明**：数值字段使用容差比较 `|expected - actual| <= 1e-6`，避免浮点运算精度问题。

## 故障排除

### 问题 1：基准文件不存在

```
pytest.skip: 基准输入文件不存在
```

**解决**：
```powershell
Copy-Item "data\BOM-COS1000334701-BA.xlsm" "tests\test_data\baseline\"
```

### 问题 2：单元格差异

```
AssertionError: 输出文件存在 15 个差异，匹配率: 99.98%
```

**排查步骤**：
1. 查看差异报告：`tests/test_data/temp_outputs/differences_*.json`
2. 确认差异是否预期（bug修复、算法改进）
3. 如果预期 → 更新基准；如果非预期 → 修复代码

### 问题 3：性能超时

```
AssertionError: 平均处理时间超过阈值: 65.23秒 > 60秒
```

**解决**：优化代码或调整 `BOMTestConfig.MAX_PROCESSING_TIME`

## 更新基准数据

代码修改**有意**改变输出时（如bug修复、算法改进）：

```powershell
# 1. 删除旧基准
Remove-Item tests\test_data\baseline\*.xlsx
Remove-Item tests\test_data\baseline\*.json

# 2. 重新生成基准
pytest tests/test_bom_processor.py::TestBOMProcessor::test_baseline_files_exist -v -s

# 3. 验证新基准
.\run_tests.ps1
```

## 文件结构

```
tests/
├── __init__.py
├── test_bom_processor.py       # 测试代码
├── pytest.ini                  # pytest配置
└── test_data/
    ├── baseline/               # 基准数据（不提交到Git）
    │   ├── BOM-COS1000334701-BA.xlsm
    │   ├── *_processed_baseline.xlsx
    │   ├── *_BOM_Sum_baseline.xlsx
    │   └── baseline_stats.json
    └── temp_outputs/           # 临时输出（不提交到Git）
        ├── differences_*.json  # 差异报告
        └── *.xlsx              # 临时测试文件
```

**注意**：`__pycache__/` 是 Python 自动生成的字节码缓存，不需要提交到 Git。

## 最佳实践

✅ 每次修改代码后运行测试  
✅ 提交代码前运行测试  
✅ 定期审查基准数据  
❌ 不要提交测试数据文件到 Git（已在 .gitignore 配置）

## 测试覆盖范围

- ✅ 工作表结构（数量、名称）
- ✅ 单元格数据（逐个验证）
- ✅ 统计数据（Level 1/2、去重等）
- ✅ 浮点数容差处理
- ✅ 性能基准
- ❌ 单元格样式（颜色、字体） - 未覆盖
- ❌ 公式验证 - 未覆盖（使用 data_only=True）

## 扩展测试

如需添加新测试用例，在 `TestBOMProcessor` 类中添加方法：

```python
def test_custom_validation(self):
    """自定义验证逻辑"""
    # 你的测试代码
    pass
```

## 联系支持

如有问题，请查看：
- 测试输出日志
- `temp_outputs/` 中的临时文件
- 处理器日志文件
