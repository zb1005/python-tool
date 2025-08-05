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
    base_df = pd.read_excel(r"E:\000000我的事项\2025-07\1131\00-审批数据模板-0729汇总v1.1.xlsx",sheet_name="计算单个节点时长")
    holiday_df = pd.read_excel(r"E:\000000我的事项\2025-07\1131\2025非工作日清单-截至Q3.xlsx", engine='openpyxl')
    special_node_df = pd.read_excel(r"E:\000000我的事项\2025-07\1131\特殊节点时长基准-250729.xlsx", engine='openpyxl')

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


    #创建流程名称与建议时长的对照关系
    node_duration_map = dict(zip(
        special_node_df['流程节点拼接'],
        special_node_df['特殊合理时长基准（天）']
    ))
    base_df['节点审批时长'] = base_df['流程节点拼接'].map(lambda x: node_duration_map.get(x, 1))
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
    base_df['该节点审批工作时长_规整'] = base_df['该节点审批工作时长'].apply(lambda x: max(0, round(x, 2)))
    base_df['节点审批时效情况：≤1；1<X≤3；>3'] = base_df['该节点审批工作时长_规整'].apply(
        lambda x: '≤1' if x <= 1 else ('1<X≤3' if 1 < x <= 3 else '>3')
    )
    base_df['节点审批延期时长(实际工作时长-节点审批时长）'] = base_df['该节点审批工作时长'] - base_df['节点审批时长']
    base_df['节点审批延期时长(实际工作时长-节点审批时长）'] = base_df['节点审批延期时长(实际工作时长-节点审批时长）'].apply(lambda x: max(0, round(x, 2)))
    base_df['延期时长情况：0；≤1；1<X≤3；>3'] = base_df['节点审批延期时长(实际工作时长-节点审批时长）'].apply(
        lambda x: '0' if x == 0 else ('≤1' if x <= 1 else ('1<X≤3' if 1 < x <= 3 else '>3'))
    )   
    

    # 保存结果到Excel（含多个sheet）
    with pd.ExcelWriter(r'E:\000000我的事项\2025-07\1131\Q3审批-节点时长统计分析.xlsx', engine='openpyxl') as writer:
        base_df.to_excel(writer, sheet_name='时长计算', index=False)
    return "Q3分析已完成，结果保存至Q3审批-节点时长统计分析.xlsx"

if __name__ == '__main__':
    result = process_approval_data()
    print(result)