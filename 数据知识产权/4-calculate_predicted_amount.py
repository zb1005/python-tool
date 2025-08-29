import pandas as pd
import os

def calculate_predicted_amount(province_folder, matrix_file, output_folder):
    # 创建输出文件夹
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    # 读取转开票重构矩阵
    matrix_df = pd.read_excel(matrix_file)
    matrix_df = matrix_df.set_index('办事处')
    
    # 定义月份列与矩阵日期的映射关系
    month_mapping = {
        '前推6个月审核金额': '2023-07-01',
        '前推5个月审核金额': '2023-08-01',
        '前推4个月审核金额': '2023-09-01',
        '前推3个月审核金额': '2023-10-01',
        '前推2个月审核金额': '2023-11-01',
        '前推1个月审核金额': '2023-12-01'
    }
    
    # 处理每个省份的文件
    for filename in os.listdir(province_folder):
        if filename.endswith('_门店金额分配.xlsx') and not filename.startswith('~$'):
            province = filename.split('_')[0]
            file_path = os.path.join(province_folder, filename)
            
            # 读取省份数据
            df = pd.read_excel(file_path)
            
            # 计算预测金额
            predicted_amounts = []
            for _, row in df.iterrows():
                office = row['办事处']
                total = 0
                
                for month_col, date_col in month_mapping.items():
                    if pd.notna(row[month_col]) and office in matrix_df.index and date_col in matrix_df.columns:
                        total += row[month_col] * matrix_df.at[office, date_col]
                
                predicted_amounts.append(round(total, 2))
            
            # 添加预测金额列
            df['预测金额'] = predicted_amounts
            
            # 保存结果
            output_path = os.path.join(output_folder, f'{province}_预测金额结果.xlsx')
            df.to_excel(output_path, index=False)
            print(f'已生成{province}预测金额结果: {output_path}')
    
    return output_folder

if __name__ == '__main__':
    # 省份合并结果文件夹路径
    province_folder = r'C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\省份合并结果'
    
    # 转开票重构矩阵文件路径
    matrix_file = r'C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\审核转开票矩阵_重构结果.xlsx'
    
    # 预测结果输出文件夹
    output_folder = r'C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\省份预测结果'
    
    # 执行计算
    calculate_predicted_amount(province_folder, matrix_file, output_folder)
    print('所有省份预测金额计算完成！')