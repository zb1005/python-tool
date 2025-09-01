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
    base_df = pd.read_excel(r"E:\000000我的事项\2025-07\1131\00-审批数据模板-0729汇总v1.1.xlsx",sheet_name="1、已批准（不含拒绝）-计算周期")
    holiday_df = pd.read_excel(r"E:\000000我的事项\2025-08\1131统计\2025非工作日清单-截至0901.xlsx", engine='openpyxl')
    # special_node_df = pd.read_excel(r"E:\000000我的事项\2025-07\1131\特殊节点时长基准-250729.xlsx", engine='openpyxl')
    print(base_df.columns)

    #对base_df建立流程编码和审批具体操作的对应关系，具体操作可能存在多个
    process_operation_map = {}
    for _, row in base_df.iterrows():
        process_code = row['流程编号']
        operation = row['审批具体操作']
        if process_code not in process_operation_map:
            process_operation_map[process_code] = []
        process_operation_map[process_code].append(operation)
    
    #找出所有审批具体操作中包含驳回，撤回，审批退回，退回，已退回的流程编码，并将所有的这部分流程编码进行剔除
    reject_process_codes = []
    for process_code, operations in process_operation_map.items():
        for operation in operations:
            if any(keyword in operation for keyword in ['驳回', '撤回', '审批退回', '退回', '已退回']):
                reject_process_codes.append(process_code)
                break
    base_df = base_df[~base_df['流程编号'].isin(reject_process_codes)]

    # #创建流程名称与建议时长的对照关系
    # node_duration_map = dict(zip(
    #     special_node_df['流程节点拼接'],
    #     special_node_df['特殊合理时长基准（天）']
    # ))
    # base_df['节点审批时长'] = base_df['流程节点拼接'].map(lambda x: node_duration_map.get(x, 1))
    # 处理假期日期
    holidays = [datetime.strptime(str(date).strip(), '%Y-%m-%d %H:%M:%S').date() for date in holiday_df['方太假期']]

    # 转换时间格式
    base_df['单个节点审批到达时间'] = pd.to_datetime(base_df['单个节点审批到达时间'], format="mixed")
    base_df['单个节点审批结束时间'] = pd.to_datetime(base_df['单个节点审批结束时间'], format="mixed")
    base_df['流程结束时间'] = pd.to_datetime(base_df['流程结束时间'], format="mixed")

    # 1. 只保留流程结束时间在2025年7月内的数据
    base_df = base_df[(base_df['流程结束时间'].dt.year == 2025) & (base_df['流程结束时间'].dt.month == 7)]

    # 2. 计算与Q2相同的时长指标
    base_df['该节点审批自然时长'] = (base_df['单个节点审批结束时间'] - base_df['单个节点审批到达时间']).dt.total_seconds() / (24 * 3600)
    base_df['该节点审批工作时长'] = base_df.apply(
        lambda row: calculate_work_duration(
            row['单个节点审批到达时间'], 
            row['单个节点审批结束时间'], 
            holidays
        ), 
        axis=1
    )

    # 3. 新建sheet：每个流程耗时（按流程编号聚合）
    process_duration = base_df.groupby('流程编号')['该节点审批工作时长'].sum().reset_index()
    process_duration.columns = ['流程编号', '流程总工作时长（天）']

    # 4. 新建sheet：每种流程耗时&频次统计（按流程名称聚合）
    # 先将流程编号和流程名称的映射关系提取出来
    process_id_name = base_df[['流程编号', '流程名称','分类']].drop_duplicates()

    process_duration = process_duration.merge(process_id_name,on='流程编号',how='left')
    # 将流程耗时数据与流程名称关联
    process_duration_with_name = process_duration.copy()
    # 按流程名称聚合计算平均工作时长和出现频次
    process_stats = process_duration_with_name.groupby('流程名称').agg(
        平均工作时长=('流程总工作时长（天）', 'mean'),
        出现频次=('流程编号', 'count')
    ).reset_index()
    process_stats['平均工作时长'] = process_stats['平均工作时长'].round(2)
    #将分类标签匹配进去
    process_stats = process_stats.merge(process_id_name[['流程名称','分类']].drop_duplicates(),on='流程名称',how='left')
    # 5. 标记耗时每个分类内部最多前三和频次最高前三
    process_stats['耗时排名'] = process_stats.groupby('分类')['平均工作时长'].rank(ascending=False, method='min').astype(int)
    process_stats['频次排名'] = process_stats.groupby('分类')['出现频次'].rank(ascending=False, method='min').astype(int)
    process_stats['耗时前三标记'] = process_stats['耗时排名'].apply(lambda x: '是' if x <= 3 else '否')
    process_stats['频次前三标记'] = process_stats['频次排名'].apply(lambda x: '是' if x <= 3 else '否')

    # 保存结果到Excel（含多个sheet）
    with pd.ExcelWriter(r'E:\000000我的事项\2025-07\1131\Q3审批结果分析1.xlsx', engine='openpyxl') as writer:
        base_df.to_excel(writer, sheet_name='原始数据', index=False)
        process_duration.to_excel(writer, sheet_name='每个流程耗时', index=False)
        process_stats.to_excel(writer, sheet_name='每种流程耗时&频次统计', index=False)

    return "Q3分析已完成，结果保存至Q3审批结果分析1.xlsx"

if __name__ == '__main__':
    result = process_approval_data()
    print(result)