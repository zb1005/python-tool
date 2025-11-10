import pandas as pd
from xlsxwriter.workbook import Workbook

def write_and_style_fixed(df, sheet_name, writer):
    """
    修复版本的write_and_style函数，解决表头位置偏移和列宽不足的问题
    
    参数:
    df: DataFrame - 要写入的数据
    sheet_name: str - 工作表名称
    writer: ExcelWriter - Excel写入器对象
    """
    
    # 获取工作簿和工作表
    workbook = writer.book
    
    # 定义表头格式
    header_format = workbook.add_format({
        'font_name': '微软雅黑',       # 字体：微软雅黑
        'font_color': 'white',         # 字体颜色：白色
        'bg_color': '#990000',         # 背景色：深红色
        'bold': True,                  # 加粗
        'align': 'center',             # 水平居中
        'valign': 'vcenter',           # 垂直居中
        'border': 1                    # 添加边框
    })
    
    # 定义数据单元格格式
    data_format = workbook.add_format({
        'font_name': '微软雅黑',        # 字体：微软雅黑
        'border': 1                    # 添加边框
    })
    
    # 先写入数据，但不包含表头（从第2行开始）
    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
    worksheet = writer.sheets[sheet_name]
    
    # 手动写入表头并应用格式 - 从第一列开始正确设置
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    # 设置数据行格式
    num_rows, num_cols = df.shape
    for row_num in range(1, num_rows + 1):  # 从第2行开始（数据行）
        for col_num in range(num_cols):
            # 获取单元格值
            cell_value = df.iloc[row_num - 1, col_num]
            # 写入数据并应用格式
            worksheet.write(row_num, col_num, cell_value, data_format)
    
    # 更精确的列宽自适应
    for col_num, col_name in enumerate(df.columns):
        # 计算列宽：取列名长度和该列数据最大长度的较大值
        max_length = max(df[col_name].astype(str).str.len().max(), len(str(col_name)))
        # 设置合适的列宽（字符数 * 1.2 + 2作为缓冲）
        column_width = max_length * 1.2 + 2
        # 限制最大列宽为50，最小为8
        column_width = min(max(column_width, 8), 50)
        worksheet.set_column(col_num, col_num, column_width)
    
    # 冻结首行以便查看表头
    worksheet.freeze_panes(1, 0)
    
    # 设置行高
    worksheet.set_row(0, 20)  # 表头行高
    for row_num in range(1, num_rows + 1):
        worksheet.set_row(row_num, 15)  # 数据行高

def create_excel_with_fixed_styles(dataframes_dict, output_path):
    """
    使用修复版本的函数创建Excel文件
    
    参数:
    dataframes_dict: dict - 字典，键为工作表名称，值为DataFrame
    output_path: str - 输出文件路径
    """
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for sheet_name, df in dataframes_dict.items():
            write_and_style_fixed(df, sheet_name, writer)
    
    print(f"Excel文件已成功创建：{output_path}")

# 使用示例
if __name__ == "__main__":
    # 示例数据
    df1 = pd.DataFrame({
        '产品线': ['油烟机产品线', '烹饪厨电产品线'],
        '数量': [100, 200],
        '占比': [0.5, 0.5]
    })
    
    df2 = pd.DataFrame({
        '年份': [2022, 2023, 2024],
        '新增数量': [50, 60, 70],
        '存活率': [0.8, 0.9, 0.95]
    })
    
    # 创建Excel文件
    dataframes = {
        '整机概览': df1,
        '存活情况': df2
    }
    
    create_excel_with_fixed_styles(dataframes, 'test_output.xlsx')