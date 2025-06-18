import pandas as pd
import numpy as np
import os
from datetime import datetime

# 处理门店数据：按省分组并保存
def process_excel(input_file, output_folder):
    # 读取Excel文件
    df = pd.read_excel(input_file, sheet_name="城市-门店")
    
    # 按省分组并保存
    for province, province_group in df.groupby('省'):
        province_file = os.path.join(output_folder, f'{province}.xlsx')
        
        # 按城市分组
        city_groups = province_group.groupby('城市')
        
        # 计算每个城市的门店数量
        city_counts = city_groups.size()
        
        # 生成600-800的随机总数
        total_stores = np.random.randint(600, 800)
        
        # 计算每个城市的新门店数量（保持比例）
        city_ratios = city_groups.size() / city_counts.sum()
        new_counts = (city_ratios * total_stores).round().astype(int)
        
        # 调整总数可能不等于total_stores的情况
        diff = total_stores - new_counts.sum()
        if diff != 0:
            new_counts.iloc[0] += diff
        
        # 创建结果DataFrame
        result = []
        for city, count in new_counts.items():
            city_group = city_groups.get_group(city)
            
            # 随机选择指定数量的门店（确保门店和简称都不重复）
            # 先按门店简称去重
            unique_abbr_group = city_group.drop_duplicates('门店简称')
            
            if len(unique_abbr_group) >= count:
                sampled = unique_abbr_group.sample(n=count, replace=False)
            else:
                # 如果唯一简称数量不足，使用所有可用的唯一简称
                sampled = unique_abbr_group.copy()
                # 记录警告信息
                print(f"警告：城市'{city}'的门店简称数量不足，无法满足需求。需要{count}个，实际只有{len(sampled)}个唯一简称。")
            
            # 添加城市和省份信息
            sampled['城市'] = city
            sampled['省'] = province
            
            result.append(sampled)
        
        # 合并并保存
        pd.concat(result).to_excel(province_file, index=False)

# 处理审核数据：按省分组并计算月度汇总
def process_audit_data(audit_file, output_folder):
    # 读取审核数据
    audit_df = pd.read_excel(audit_file, sheet_name="只要下半年")
    
    # 按省分组处理审核数据
    for province, province_group in audit_df.groupby('省'):
        # 按办事处(即城市)和月份汇总金额，将办事处重命名为城市
        monthly_summary = province_group.groupby(['办事处', province_group['日期'].dt.to_period('M')])['金额'].sum().unstack()
        monthly_summary = monthly_summary.reset_index().rename(columns={'办事处': '城市'})
        
        # 重命名月份列为前推审核金额
        monthly_summary = monthly_summary.rename(columns={
            '2023-07': '前推6个月审核金额',
            '2023-08': '前推5个月审核金额',
            '2023-09': '前推4个月审核金额',
            '2023-10': '前推3个月审核金额',
            '2023-11': '前推2个月审核金额',
            '2023-12': '前推1个月审核金额'
        })
        
        # 保存为省份Excel文件
        output_file = os.path.join(output_folder, f'{province}_审核汇总.xlsx')
        monthly_summary.to_excel(output_file)

import numpy as np

def assign_store_levels(store_df):
    # 按门店规模分配等级1-10
    # 生成正态分布数据实现两头小中间大的分布
    np.random.seed(42)
    normal_data = np.random.normal(loc=5.5, scale=2, size=len(store_df))
    store_df['门店等级'] = pd.qcut(normal_data, 10, labels=range(1, 11)).astype(int)
    # 计算基础权重
    # 使用平方级差拉大高低等级差距
    store_df['基础权重'] = store_df['门店等级'] **2 / store_df['门店等级'].pow(2).sum()
    return store_df

def split_amounts_by_level(store_df, amount_df):
    # 复制原始数据
    result_df = store_df.copy()
    # 获取所有月份列
    month_columns = [col for col in amount_df.columns if col != '城市']
    
    # 按办事处和城市分组处理
    for city, office_group in amount_df.groupby('城市'):
        # 筛选当前城市的门店
        city_stores = result_df[result_df['城市'] == city].copy()
        if city_stores.empty:
            continue
        
        for month in month_columns:
            total_amount = office_group.iloc[0][month]
            if pd.isna(total_amount):
                result_df.loc[city_stores.index, month] = 0
                continue
            
            # 生成随机扰动因子(±5%)
            np.random.seed(42)  # 固定随机种子确保可复现
            random_factors = np.random.uniform(0.95, 1.15, len(city_stores))
            
            # 计算带扰动的权重
            weighted = city_stores['基础权重'] * random_factors
            normalized_weights = weighted / weighted.sum()
            
            # 分配金额
            amounts = (normalized_weights * total_amount).round(2)
            result_df.loc[city_stores.index, month] = amounts
            
            # 确保总和与原始金额一致
            diff = total_amount - amounts.sum()
            if abs(diff) > 0.001:
                # 调整最大金额的门店
                max_idx = amounts.idxmax()
                result_df.at[max_idx, month] += diff
    
    return result_df

def combine_province_data(store_folder, audit_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    # 获取所有省份文件
    store_files = {os.path.splitext(f)[0].split('_')[0]: f for f in os.listdir(store_folder) if f.endswith('.xlsx')}
    audit_files = {f.split('_审核汇总')[0]: f for f in os.listdir(audit_folder) if f.endswith('.xlsx') and '_审核汇总' in f}
    print(store_files.keys())
    print(audit_files.keys())
    # 处理每个省份
    for province in store_files.keys() & audit_files.keys():
        # 读取数据
        store_path = os.path.join(store_folder, store_files[province])
        audit_path = os.path.join(audit_folder, audit_files[province])
        
        store_df = pd.read_excel(store_path)
        audit_df = pd.read_excel(audit_path)
        
        # 分配门店等级
        store_df = assign_store_levels(store_df)
        
        # 拆分金额
        result_df = split_amounts_by_level(store_df, audit_df)
        
        # 重命名月份列
        month_rename = {
            '2023-07': '前推6个月审核金额',
            '2023-08': '前推5个月审核金额',
            '2023-09': '前推4个月审核金额',
            '2023-10': '前推3个月审核金额',
            '2023-11': '前推2个月审核金额',
            '2023-12': '前推1个月审核金额'
        }
        result_df = result_df.rename(columns=month_rename)
        
        # 重命名城市为办事处并添加城市名称列
        result_df = result_df.rename(columns={'城市': '办事处'})
        result_df['城市名称'] = result_df['办事处']
        
        # 调整列顺序
        column_order = [
            '省', '城市名称', '办事处', '门店编码', '门店简称',
            '前推6个月审核金额', '前推5个月审核金额', '前推4个月审核金额',
            '前推3个月审核金额', '前推2个月审核金额', '前推1个月审核金额'
        ]
        result_df = result_df[column_order]
        
        # 保存结果
        output_path = os.path.join(output_folder, f'{province}_门店金额分配.xlsx')
        result_df.to_excel(output_path, index=False)
    
    return output_folder

# 使用示例
if __name__ == '__main__':
    # 处理门店数据
    input_file = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\门店信息.XLSX"
    store_output_folder = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\门店"
    if not os.path.exists(store_output_folder):
        os.makedirs(store_output_folder)
    process_excel(input_file, store_output_folder)
    
    # 处理审核数据
    audit_file = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\审核订单_时间预测数据.xlsx"
    audit_output_folder = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\金额汇总"
    if not os.path.exists(audit_output_folder):
        os.makedirs(audit_output_folder)
    process_audit_data(audit_file, audit_output_folder)
    
    # 合并省份数据
    combined_output_folder = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\总\省份合并结果"
    combine_province_data(store_output_folder, audit_output_folder, combined_output_folder)
    print(f"省份数据合并完成，结果保存在: {combined_output_folder}")
