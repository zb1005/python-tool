import pandas as pd
from datetime import datetime, timedelta

def is_workday(date, holidays):
    # 判断是否为工作日（非周末且非节假日）
    return date not in holidays

def calculate_work_duration(start_time, end_time, holidays):
    # 计算工作时间（排除节假日和周末）
    current = start_time
    work_duration = timedelta()
    
    while current < end_time:
        next_day = (current + timedelta(days=1)).replace(hour=0, minute=0, second=0)
        if is_workday(current.date(), holidays):
            work_duration += min(next_day, end_time) - current
        current = next_day
    
    return work_duration.total_seconds() / (24 * 3600)  # 转换为天数


def process_approval_data():
    # 读取Q3相关Excel文件（路径需根据实际Q3数据调整）
    base_df = pd.read_excel(r"E:\000000我的事项\2025-07\1131\00-审批数据模板-0729汇总v1.1.xlsx",sheet_name="2、已批准（不含加签）-计算节点")


    # 转换时间格式
    base_df['单个节点审批到达时间'] = pd.to_datetime(base_df['单个节点审批到达时间'], format="mixed")
    base_df['单个节点审批结束时间'] = pd.to_datetime(base_df['单个节点审批结束时间'], format="mixed")
    base_df['流程结束时间'] = pd.to_datetime(base_df['流程结束时间'], format="mixed")

    # 1. 只保留流程结束时间在2025年7月内的数据
    base_df = base_df[(base_df['流程结束时间'].dt.year == 2025) & (base_df['流程结束时间'].dt.month == 7)]

    # 3. 新建sheet：每个流程耗时（按流程编号聚合）
    process_duration = base_df.groupby('流程编号')['单个节点审批到达时间'].nunique().reset_index()
    process_duration.columns = ['流程编号', '流程总节点数']

    # 4. 新建sheet：每种流程耗时&频次统计（按流程名称聚合）
    # 先将流程编号和流程名称的映射关系提取出来
    process_id_name = base_df[['流程编号', '流程名称','分类']].drop_duplicates()
    # 将流程耗时数据与流程名称关联
    process_duration_with_name = process_duration.merge(process_id_name, on='流程编号')

    # 对于节点数大于5的流程进行标记
    process_duration_with_name['节点数大于5'] = process_duration_with_name['流程总节点数'].apply(lambda x: 1 if x > 5 else 0)
    
    # 按流程名称聚合计算平均节点数和出现频次，以及计算节点数大于5的比例
    process_stats = process_duration_with_name.groupby('流程名称').agg(
        平均节点数=('流程总节点数', 'mean'),
        出现频次=('流程编号', 'count'),
        节点数大于5比例=('节点数大于5', lambda x: x.sum() / len(x))  # 直接计算比例
    ).reset_index()
    process_stats = process_stats.merge(process_id_name[['流程名称','分类']].drop_duplicates(),on='流程名称')

    # 保存结果到Excel（含多个sheet）
    with pd.ExcelWriter(r'E:\000000我的事项\2025-07\1131\Q3审批结果分析2.xlsx', engine='openpyxl') as writer:
        process_duration_with_name.to_excel(writer, sheet_name='每个流程节点数', index=False)
        process_stats.to_excel(writer, sheet_name='每种流程节点统计', index=False)

    return "Q3分析已完成，结果保存至Q3审批结果分析2.xlsx"

if __name__ == '__main__':
    result = process_approval_data()
    print(result)