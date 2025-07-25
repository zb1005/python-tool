import pandas as pd
import os

def process_invoice_matrix(input_file, sheet_name, output_file):
    # 读取Excel文件
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    
    # 定义时间映射关系: T+0对应2023-12, T+1对应2023-11, ..., T+5对应2023-07
    time_mapping = {
        'T+0': '2023-12-01',
        'T+1': '2023-11-01',
        'T+2': '2023-10-01',
        'T+3': '2023-09-01',
        'T+4': '2023-08-01',
        'T+5': '2023-07-01'
    }
    
    # 获取所有日期列
    date_columns = [col for col in df.columns if col not in ['办事处', '时间']]
    
    # 创建结果列表
    result = []
    
    # 按办事处分组处理
    for office, office_group in df.groupby('办事处'):
        office_data = {'办事处': office}
        
        # 处理每个时间点
        for t_value, date_col in time_mapping.items():
            # 检查日期列是否存在
            if date_col not in date_columns:
                print(f"警告: 办事处'{office}'的日期列'{date_col}'不存在")
                office_data[date_col] = None
                continue
            
            # 提取对应时间点的数据
            t_row = office_group[office_group['时间'] == t_value]
            if not t_row.empty:
                office_data[date_col] = t_row.iloc[0][date_col]
            else:
                print(f"警告: 办事处'{office}'的时间'{t_value}'不存在")
                office_data[date_col] = None
        
        result.append(office_data)
    
    # 转换为DataFrame并按日期排序列
    result_df = pd.DataFrame(result)
    sorted_columns = ['办事处'] + sorted(date_columns, key=lambda x: pd.to_datetime(x))
    result_df = result_df[sorted_columns]
    
    # 保存结果
    result_df.to_excel(output_file, index=False)
    print(f"处理完成，结果保存在: {output_file}")
    return result_df

if __name__ == '__main__':
    # 输入文件路径
    input_file = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\审核转开票矩阵_2025_f.xlsx"
    sheet_name = "只要下半年"
    
    # 输出文件路径
    output_dir = os.path.dirname(input_file)
    output_file = os.path.join(output_dir, "审核转开票矩阵_重构结果.xlsx")
    
    # 处理数据
    process_invoice_matrix(input_file, sheet_name, output_file)