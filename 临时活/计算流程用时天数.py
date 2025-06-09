import pandas as pd
import datetime
from datetime_utils import calculate_work_hours  # 导入之前写的计算函数
from datetime_utils import calculate_work_days  # 导入之前写的计算函数

def process_excel_file(input_path, output_path):
    # 读取Excel文件（支持xlsx和xls格式）
    df = pd.read_excel(input_path,sheet_name='流程&批导')
    print("===========",df)
    
    # 转换时间列（使用实际列名）
    df['创建日期'] = pd.to_datetime(df['创建日期'], errors='coerce')
    print("===========",df['创建日期'])
    df['采购审核日期'] = pd.to_datetime(df['采购审核日期'], errors='coerce')
    
    # 新增流程结束时间列转换（假设存在该列）
    df['流程结束日期'] = pd.to_datetime(df['流程结束日期'], errors='coerce')
    
    # 新增有效工时列
    df['有效天数-采购&主数据'] = df.apply(
        lambda row: calculate_work_days(row['创建日期'], row['流程结束日期']) ,
        axis=1
    )
    df['有效天数-采购'] = df.apply(
    lambda row: calculate_work_days(row['创建日期'], row['采购审核日期']) ,
    axis=1
    )
        # 新增有效工时列
    df['有效天数-主数据'] = df.apply(
        lambda row: calculate_work_days(row['采购审核日期'], row['流程结束日期']) ,
        axis=1
    )
    df['有效天数-NPI'] = df.apply(
        lambda row: calculate_work_days(row['物料创建日期'], row['创建日期']) ,
        axis=1
    )
    df['有效天数-总流程'] = df.apply(
    lambda row: calculate_work_days(row['物料创建日期'], row['流程结束日期']) ,
    axis=1
    )

    # 新增有效小时数列
    df['有效小时数-NPI'] = df.apply(
        lambda row: calculate_work_hours(row['物料创建日期'], row['创建日期']),
        axis=1
    )
    df['有效小时数-采购'] = df.apply(
        lambda row: calculate_work_hours(row['创建日期'], row['采购审核日期']),
        axis=1
    )
    df['有效小时数-主数据'] = df.apply(
        lambda row: calculate_work_hours(row['采购审核日期'], row['流程结束日期']),
        axis=1
    )
    df['有效小时数-总流程'] = df.apply(
        lambda row: calculate_work_hours(row['物料创建日期'], row['流程结束日期']),
        axis=1
    )

    #保留物料创建日期在2024年11月1日到2025年3月31日之间的记录
    df = df[(df['物料创建日期'] >= '2024-11-01') & (df['物料创建日期'] <= '2025-03-31')]

    #在同一个物料编码下，只保留创建日期最早的记录
    df = df.sort_values(by=['物料编码', '创建日期']).drop_duplicates(subset=['物料编码'], keep='first')
    
    # # 处理空值情况
    # if df['有效时长'].isnull().any():
    #     print("警告：存在无法解析的时间格式，已自动替换为0")
    #     df['有效时长'].fillna(0, inplace=True)
    
    # 保存结果到新文件
    df.to_excel(output_path, index=False)
    print(f"处理完成，结果已保存至：{output_path}")

# 使用示例（请替换实际文件路径）
if __name__ == "__main__":
    input_file = r"C:\Users\zhangbon\Desktop\流程记录.XLSX"  # 原始Excel路径
    output_file = r"C:\Users\zhangbon\Desktop\流程记录-总计.xlsx"  # 输出路径
    
    process_excel_file(input_file, output_file)