# 修复后的Excel写入代码 - 直接替换到您的Jupyter Notebook中

# 替换原来的Excel写入代码段为以下代码：

with pd.ExcelWriter(fr"D:\000物料报表\物料运维报告\物料运维报告.xlsx", engine='xlsxwriter') as writer:
    workbook = writer.book
    
    # 表头格式
    header_format = workbook.add_format({
        'font_name': '微软雅黑',
        'font_color': 'white',
        'bg_color': '#990000',
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # 数据格式
    data_format = workbook.add_format({
        'font_name': '微软雅黑',
        'border': 1
    })
    
    def write_and_style_fixed(df, sheet_name):
        """修复版本的write_and_style函数"""
        # 创建新的工作表
        worksheet = workbook.add_worksheet(sheet_name)
        
        # 手动写入表头并应用格式 - 从第一列开始
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # 写入数据并应用格式
        for row_num, row_data in enumerate(df.values, start=1):
            for col_num, cell_value in enumerate(row_data):
                worksheet.write(row_num, col_num, cell_value, data_format)
        
        # 精确计算列宽
        for col_num, col_name in enumerate(df.columns):
            # 计算列名长度和数据最大长度
            header_length = len(str(col_name))
            data_max_length = df[col_name].astype(str).str.len().max()
            max_length = max(header_length, data_max_length)
            
            # 设置合适的列宽（字符数 * 1.1 + 1.5作为缓冲）
            column_width = max_length * 1.1 + 1.5
            # 限制最小8，最大40
            column_width = min(max(column_width, 8), 40)
            worksheet.set_column(col_num, col_num, column_width)
        
        # 冻结首行
        worksheet.freeze_panes(1, 0)
        # 设置表头行高
        worksheet.set_row(0, 20)
    
    # 写入各个工作表
    write_and_style_fixed(df1_with_total, '整机概览')
    write_and_style_fixed(df2_with_total, '各产品线整机存活情况')
    write_and_style_fixed(df3, '在产型号生产采购订单情况')
    write_and_style_fixed(df4, '淘汰阶段型号采购订单情况')
    write_and_style_fixed(df5, '零件冻结&自制外购情况')
    write_and_style_fixed(df6, '零部件新增情况')
    write_and_style_fixed(df7, '新增零部件使用情况')

print("Excel文件已成功创建：物料运维报告.xlsx")