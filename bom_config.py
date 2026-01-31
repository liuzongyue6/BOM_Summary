"""
BOM Processor Configuration
集中管理所有列配置和处理参数
"""

class BOMConfig:
    """BOM处理配置类"""
    
    # ============ 必需列配置 ============
    # 用于验证输入文件是否包含必需的列
    REQUIRED_COLUMNS = [
        'BOM Line',
        'CAD OEM Part Number',
        'CAD OEM Rev',
        'Quantity'
    ]
    
    # ============ 压缩表输出列配置 ============
    # 定义压缩表(Compress sheets)要输出的列
    # 注意：这些列会用于去重和聚合
    COMPRESS_OUTPUT_COLUMNS = [
        'CAD OEM Part Number',
        'CAD OEM Rev',
        'Material Spec',
        'CAD Oem Name',
        'Thickness',
        'Weight',
        'Area',
        'Quantity'
    ]
    
    # ============ 去重分组键配置 ============
    # 用于去重的关键字段（相同的part number和revision视为同一零件）
    GROUPING_KEY_COLUMNS = [
        'CAD OEM Part Number',
        'CAD OEM Rev'
    ]
    
    # ============ 元数据列配置 ============
    # 在去重时需要保留的元数据字段（非累加字段）
    METADATA_COLUMNS = [
        'Material Spec',
        'CAD Oem Name',
        'Thickness',
        'Weight',
        'Area'
    ]
    
    # ============ 累加列配置 ============
    # 需要累加的字段
    ACCUMULATION_COLUMN = 'Quantity'
    
    # ============ BOM汇总表输出列配置 ============
    # BOM_Summary sheet的固定列（在动态列之前）
    SUMMARY_FIXED_COLUMNS = [
        'CAD OEM Part Number',
        'CAD OEM Rev',
        'Material Spec',
        'CAD Oem Name',
        'Thickness',
        'Weight',
        'Area'
    ]
    
    # ============ BOM层级标识列 ============
    BOM_LINE_COLUMN = 'BOM Line'
    
    # ============ 默认值配置 ============
    DEFAULT_QUANTITY = 1  # 当Quantity为空或0时的默认值
    
    @classmethod
    def get_compress_column_indices(cls):
        """获取压缩表列的索引映射 {列名: 列号(1-based)}"""
        return {col: idx + 1 for idx, col in enumerate(cls.COMPRESS_OUTPUT_COLUMNS)}
    
    @classmethod
    def get_metadata_dict(cls):
        """获取初始化的元数据字典"""
        metadata = {col: None for col in cls.METADATA_COLUMNS}
        metadata[cls.ACCUMULATION_COLUMN] = 0
        return metadata
    
    @classmethod
    def validate_columns(cls, headers: list) -> tuple:
        """
        验证必需列是否存在
        
        Args:
            headers: 列名列表
            
        Returns:
            (is_valid, missing_columns)
        """
        missing = [col for col in cls.REQUIRED_COLUMNS if col not in headers]
        return len(missing) == 0, missing


# ============ 向后兼容：提供模块级别的访问 ============
REQUIRED_COLUMNS = BOMConfig.REQUIRED_COLUMNS
COMPRESS_COLUMNS = BOMConfig.COMPRESS_OUTPUT_COLUMNS
GROUPING_KEYS = BOMConfig.GROUPING_KEY_COLUMNS
METADATA_COLUMNS = BOMConfig.METADATA_COLUMNS