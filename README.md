# 数据处理与文档生成系统

## 系统功能

本系统是一个封装了 pandas、xlsxwriter、openpyxl 和 python-docx 的数据处理与文档生成工具，主要功能包括：

1. **数据处理**：接收外部数据库查到的字典数据，转换为 DataFrame 进行处理
2. **农历日期计算**：支持将公历日期转换为农历日期，用于同比分析
3. **数据计算**：
   - 求和、最大值、最小值、平均值
   - 环比分析
   - 同比分析（支持公历和农历日期）
   - 占比计算
   - 未来预估（按比值同比历史数据预估）
4. **Excel 表格生成**：
   - 自动创建 Excel 文件并填充数据
   - 支持设置表头格式（加粗、居中、背景色等）
   - 支持设置数据格式（数值保留两位小数，百分比保留两位小数）
   - 自动调整列宽
5. **Word 文档生成**：
   - 支持填充 Word 文档模板
   - 支持从零生成 Word 文档
   - 支持设置文档格式（标题居中、加粗等）
6. **文件路径处理**：
   - 自动创建不存在的目录
   - 自动重命名已存在的文件

## 安装依赖

系统需要安装以下依赖库：

```bash
pip install pandas xlsxwriter openpyxl python-docx lunar-python
```

## 使用方法

### 1. 导入 DataProcessor 类

```python
from data_processor import DataProcessor
```

### 2. 创建 DataProcessor 实例

```python
processor = DataProcessor()
```

### 3. 处理数据

```python
# 示例数据（模拟从数据库查询到的字典数据）
data = [
    {'date': '2023-01-01', 'value': 100, 'category': 'Category 1'},
    {'date': '2023-02-01', 'value': 200, 'category': 'Category 2'},
    # 更多数据...
]

# 处理数据
df = processor.process_data(data)
```

### 4. 计算指标

```python
# 使用公历日期计算指标
result = processor.calculate_metrics(df, 'date', 'value', use_lunar=False)

# 使用农历日期计算指标
result_lunar = processor.calculate_metrics(df, 'date', 'value', use_lunar=True)

# 获取计算后的数据和指标
data_with_metrics = result['data']
metrics = result['metrics']
```

### 5. 生成 Excel 表格

```python
# 生成 Excel 表格
excel_path = processor.create_excel(data_with_metrics, 'output/data.xlsx', sheet_name='Metrics', title='数据指标分析')
print(f"Excel 表格生成成功: {excel_path}")
```

### 6. 生成 Word 文档

```python
# 从零生成 Word 文档
word_path = processor.generate_word_from_scratch(data_with_metrics, 'output/report.docx', '数据指标分析报告')
print(f"Word 文档生成成功: {word_path}")

# 填充 Word 文档模板（需要先创建模板文件）
template_path = 'template.docx'  # 包含占位符如 {total}, {max}, {min} 等
template_data = {
    'total': metrics['total'],
    'max': metrics['max'],
    'min': metrics['min'],
    'average': metrics['average'],
    'estimated_value': metrics['estimated_value']
}
word_template_path = processor.fill_word_template(template_path, template_data, 'output/template_report.docx')
print(f"Word 文档模板填充成功: {word_template_path}")
```

## 示例代码

完整的示例代码请参考 [example.py](file:///workspace/example.py) 文件。

## 注意事项

1. **数据格式**：输入数据必须是字典列表或单个字典，包含日期字段和值字段
2. **日期格式**：日期字段必须是可转换为 datetime 的字符串格式，如 '2023-01-01'
3. **文件路径**：输出文件路径可以是相对路径或绝对路径，系统会自动创建不存在的目录
4. **文件重命名**：如果输出文件已存在，系统会自动在文件名后添加数字后缀进行重命名
5. **数值格式化**：系统会自动对数值保留两位小数，对百分比保留两位小数（四位小数）
6. **农历日期**：使用 lunar-python 库进行农历日期转换，确保已安装该库

## 示例输出

运行示例代码后，系统会生成以下文件：

1. **Excel 表格**：包含原始数据和计算的指标，格式美观，包含表头加粗、居中、自动调整列宽等设置
2. **Word 文档**：包含数据表格和标题，格式美观，包含标题居中、加粗等设置
3. **嵌套目录**：系统会自动创建嵌套目录结构
4. **自动重命名**：如果文件已存在，系统会自动重命名

## 总结

本系统提供了一个完整的数据处理与文档生成解决方案，支持从数据处理、指标计算到 Excel 表格和 Word 文档生成的全流程。系统具有良好的可扩展性和易用性，可以根据实际需求进行定制和扩展。