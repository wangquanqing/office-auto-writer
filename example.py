from data_processor import DataProcessor
from datetime import datetime, timedelta
import random

# 创建 DataProcessor 实例
processor = DataProcessor()

# 生成模拟数据（模拟从数据库查询到的字典数据）
def generate_sample_data():
    data = []
    start_date = datetime(2023, 1, 1)
    for i in range(12):
        date = start_date + timedelta(days=i*30)
        data.append({
            'date': date.strftime('%Y-%m-%d'),
            'value': random.randint(100, 1000),
            'category': f'Category {i%3 + 1}'
        })
    return data

# 生成模拟数据
sample_data = generate_sample_data()
print("生成的模拟数据:")
for item in sample_data:
    print(item)

# 处理数据
df = processor.process_data(sample_data)
print("\n处理后的数据:")
print(df)

# 计算指标（使用公历日期）
result = processor.calculate_metrics(df, 'date', 'value', use_lunar=False)
print("\n计算的指标:")
print(result['metrics'])
print("\n计算后的数据:")
print(result['data'])

# 计算指标（使用农历日期）
result_lunar = processor.calculate_metrics(df, 'date', 'value', use_lunar=True)
print("\n使用农历日期计算的指标:")
print(result_lunar['metrics'])
print("\n使用农历日期计算后的数据:")
print(result_lunar['data'])

# 生成 Excel 表格
excel_path = processor.create_excel(result['data'], 'output/data.xlsx', sheet_name='Metrics', title='数据指标分析')
print(f"\nExcel 表格生成成功: {excel_path}")

# 生成 Word 文档（从零开始）
word_path = processor.generate_word_from_scratch(result['data'], 'output/report.docx', '数据指标分析报告')
print(f"\nWord 文档生成成功: {word_path}")

# 测试文件路径处理和自动重命名功能
print("\n测试文件路径处理和自动重命名功能:")
# 再次生成相同路径的文件，应该会自动重命名
excel_path2 = processor.create_excel(result['data'], 'output/data.xlsx', sheet_name='Metrics', title='数据指标分析')
print(f"第二次生成 Excel 表格: {excel_path2}")

word_path2 = processor.generate_word_from_scratch(result['data'], 'output/report.docx', '数据指标分析报告')
print(f"第二次生成 Word 文档: {word_path2}")

# 测试嵌套目录
nested_excel_path = processor.create_excel(result['data'], 'output/nested/folder/data.xlsx', sheet_name='Metrics', title='数据指标分析')
print(f"\n生成到嵌套目录的 Excel 表格: {nested_excel_path}")

nested_word_path = processor.generate_word_from_scratch(result['data'], 'output/nested/folder/report.docx', '数据指标分析报告')
print(f"生成到嵌套目录的 Word 文档: {nested_word_path}")

print("\n所有测试完成!")