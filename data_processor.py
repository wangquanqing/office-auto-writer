import os
import re
from datetime import datetime
import pandas as pd
from xlsxwriter import Workbook
from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lunar_python import Solar

class DataProcessor:
    def __init__(self):
        pass
    
    def process_data(self, data):
        """处理输入数据，转换为DataFrame"""
        if isinstance(data, list) and all(isinstance(item, dict) for item in data):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            return pd.DataFrame([data])
        else:
            raise ValueError("数据格式错误，应为字典列表或单个字典")
    
    def solar_to_lunar(self, date_str):
        """将公历日期转换为农历日期"""
        try:
            date = pd.to_datetime(date_str)
            solar = Solar.fromDate(date)
            lunar = solar.getLunar()
            # 返回农历年月信息，格式：年-月-日
            return f"{lunar.getYear()}-{lunar.getMonth()}-{lunar.getDay()}"
        except Exception as e:
            print(f"日期转换失败: {e}")
            return date_str
    
    def calculate_metrics(self, df, date_column, value_column, use_lunar=False):
        """计算各种指标"""
        # 确保日期列存在
        if date_column not in df.columns:
            raise ValueError(f"日期列 '{date_column}' 不存在")
        # 确保值列存在
        if value_column not in df.columns:
            raise ValueError(f"值列 '{value_column}' 不存在")
        
        # 转换日期格式
        df[date_column] = pd.to_datetime(df[date_column])
        # 按日期排序
        df = df.sort_values(by=date_column)
        
        # 计算求和
        total = df[value_column].sum()
        # 计算最大值
        max_value = df[value_column].max()
        # 计算最小值
        min_value = df[value_column].min()
        # 计算平均值
        average = df[value_column].mean()
        
        # 计算环比
        df['环比'] = df[value_column].pct_change()
        
        # 计算同比
        if use_lunar:
            # 使用农历日期进行同比分析
            df['农历日期'] = df[date_column].apply(self.solar_to_lunar)
            # 这里简化处理，实际应用中可能需要更复杂的逻辑
            # 假设数据是按日期顺序排列的，且每年对应位置的数据是同比数据
            df['同比'] = df[value_column].pct_change(periods=365)  # 假设是日度数据
        else:
            # 使用公历日期进行同比分析
            df['同比'] = df[value_column].pct_change(periods=12)  # 假设是月度数据
        
        # 计算占比
        df['占比'] = df[value_column] / total
        
        # 计算未来预估（按比值同比历史数据预估）
        if len(df) > 0:
            # 计算历史数据的平均同比增长率，排除 NaN 值
            historical_growth_rate = df['同比'].dropna().mean()
            # 预估未来值
            last_value = df[value_column].iloc[-1]
            if pd.isna(historical_growth_rate):
                # 如果没有有效的同比数据，使用平均值作为预估
                estimated_value = df[value_column].mean()
            else:
                estimated_value = last_value * (1 + historical_growth_rate)
        else:
            estimated_value = 0
        
        return {
            'data': df,
            'metrics': {
                'total': total,
                'max': max_value,
                'min': min_value,
                'average': average,
                'estimated_value': estimated_value
            }
        }
    
    def create_excel(self, data, output_path, sheet_name='Sheet1', title=None):
        """创建Excel文件并填充数据"""
        # 确保输出目录存在
        self._ensure_directory(output_path)
        # 处理文件重命名
        output_path = self._handle_file_exists(output_path)
        
        # 创建Excel文件，添加 nan_inf_to_errors 选项以处理 NaN 值
        workbook = Workbook(output_path, {'nan_inf_to_errors': True})
        worksheet = workbook.add_worksheet(sheet_name)
        
        # 设置格式
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#E0E0E0',
            'border': 1
        })
        
        data_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # 写入标题
        if title:
            worksheet.merge_range(0, 0, 0, len(data.columns) - 1, title, header_format)
            start_row = 1
        else:
            start_row = 0
        
        # 写入表头
        for col, header in enumerate(data.columns):
            worksheet.write(start_row, col, header, header_format)
        
        # 写入数据
        for row, (_, row_data) in enumerate(data.iterrows(), start=start_row + 1):
            for col, value in enumerate(row_data):
                # 格式化数值
                if isinstance(value, (int, float)):
                    # 检查是否为百分比
                    if '环比' in data.columns[col] or '同比' in data.columns[col] or '占比' in data.columns[col]:
                        # 百分比保留两位小数（四位小数）
                        worksheet.write(row, col, value, workbook.add_format({
                            'align': 'center',
                            'valign': 'vcenter',
                            'border': 1,
                            'num_format': '0.00%'
                        }))
                    else:
                        # 数值保留两位小数
                        worksheet.write(row, col, value, workbook.add_format({
                            'align': 'center',
                            'valign': 'vcenter',
                            'border': 1,
                            'num_format': '0.00'
                        }))
                else:
                    worksheet.write(row, col, value, data_format)
        
        # 自动调整列宽
        for col in range(len(data.columns)):
            worksheet.set_column(col, col, 15)
        
        workbook.close()
        return output_path
    
    def fill_word_template(self, template_path, data, output_path):
        """填充Word文档模板"""
        # 确保输出目录存在
        self._ensure_directory(output_path)
        # 处理文件重命名
        output_path = self._handle_file_exists(output_path)
        
        # 加载模板
        doc = Document(template_path)
        
        # 填充数据
        for paragraph in doc.paragraphs:
            for key, value in data.items():
                placeholder = f"{{{key}}}"
                if placeholder in paragraph.text:
                    # 替换文本
                    paragraph.text = paragraph.text.replace(placeholder, str(value))
                    # 设置格式
                    for run in paragraph.runs:
                        run.font.size = Pt(12)
        
        # 保存文档
        doc.save(output_path)
        return output_path
    
    def generate_word_from_scratch(self, data, output_path, title):
        """从零生成Word文档"""
        # 确保输出目录存在
        self._ensure_directory(output_path)
        # 处理文件重命名
        output_path = self._handle_file_exists(output_path)
        
        # 创建文档
        doc = Document()
        
        # 添加标题
        title_paragraph = doc.add_heading(title, 0)
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 添加表格
        if isinstance(data, pd.DataFrame):
            table = doc.add_table(rows=1, cols=len(data.columns))
            hdr_cells = table.rows[0].cells
            
            # 设置表头
            for i, col in enumerate(data.columns):
                hdr_cells[i].text = col
                # 表头加粗
                for paragraph in hdr_cells[i].paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                        run.font.size = Pt(11)
            
            # 填充数据
            for _, row_data in data.iterrows():
                row_cells = table.add_row().cells
                for i, value in enumerate(row_data):
                    # 格式化数值
                    if isinstance(value, (int, float)):
                        # 检查是否为百分比
                        if '环比' in data.columns[i] or '同比' in data.columns[i] or '占比' in data.columns[i]:
                            # 百分比保留两位小数（四位小数）
                            row_cells[i].text = f"{value:.2%}"
                        else:
                            # 数值保留两位小数
                            row_cells[i].text = f"{value:.2f}"
                    else:
                        row_cells[i].text = str(value)
        
        # 保存文档
        doc.save(output_path)
        return output_path
    
    def _ensure_directory(self, file_path):
        """确保目录存在"""
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
    
    def _handle_file_exists(self, file_path):
        """处理文件已存在的情况"""
        if os.path.exists(file_path):
            base, ext = os.path.splitext(file_path)
            counter = 1
            new_path = f"{base}_{counter}{ext}"
            while os.path.exists(new_path):
                counter += 1
                new_path = f"{base}_{counter}{ext}"
            return new_path
        return file_path