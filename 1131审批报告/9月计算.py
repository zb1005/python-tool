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
    # 读取四个Excel文件
    base_df = pd.read_excel(r"E:\000000我的事项\2025-08\1131统计\00-8月当月审批节点数据-0901.xlsx", engine='openpyxl')
    holiday_df = pd.read_excel(r"E:\000000我的事项\2025-08\1131统计\2025非工作日清单-截至0901.xlsx", engine='openpyxl')
    special_node_df = pd.read_excel(r"E:\000000我的事项\2025-08\1131统计\特殊节点时长基准-250901.xlsx",sheet_name='00', engine='openpyxl')
    

    # 3. 处理假期日期
    holidays = [datetime.strptime(str(date).strip(), '%Y-%m-%d %H:%M:%S').date() for date in holiday_df['方太假期']]
    
    # 4. 计算自然时长和工作时长
    base_df['单个节点审批到达时间'] = pd.to_datetime(base_df['单个节点审批到达时间'], format="mixed")
    base_df['单个节点审批结束时间'] = pd.to_datetime(base_df['单个节点审批结束时间'], format="mixed")
    
    base_df['该节点审批自然时长'] = (base_df['单个节点审批结束时间'] - base_df['单个节点审批到达时间']).dt.total_seconds() / (24 * 3600)
    
    base_df['该节点审批工作时长'] = base_df.apply(
        lambda row: calculate_work_duration(
            row['单个节点审批到达时间'], 
            row['单个节点审批结束时间'], 
            holidays
        ), 
        axis=1
    )
    
    # 5. 规整工作时长（保留2位小数），小于0时设为0
    base_df['该节点审批工作时长_规整'] = base_df['该节点审批工作时长'].apply(lambda x: max(0, round(x, 2)))
    
    #6.判断节点审批时效情况，按照1，2，3天分为3级
    base_df['节点审批时效情况：≤1；1<X≤3；>3'] = base_df['该节点审批工作时长_规整'].apply(
        lambda x: '≤1' if x <= 1 else ('1<X≤3' if 1 < x <= 3 else  '>3')
    )

    # 7. 创建流程名称与建议时长的对照关系，依据特殊节点时长基准表
    node_duration_map = dict(zip(
        special_node_df['流程节点拼接'],
        special_node_df['特殊合理时长基准（天）']
    ))
    
    # 8. 匹配节点合理审批时长，默认值为1
    base_df['节点审批时长'] = base_df['流程节点拼接'].map(lambda x: node_duration_map.get(x, 1))

    # 9. 计算节点审批延期时长
    base_df['节点审批延期时长(实际工作时长-节点审批时长）'] = base_df['该节点审批工作时长'] - base_df['节点审批时长']
    base_df['节点审批延期时长(实际工作时长-节点审批时长）'] = base_df['节点审批延期时长(实际工作时长-节点审批时长）'].apply(lambda x: max(0, round(x, 2)))
    # 10. 判断节点审批延期时长情况，按照1，2，3天分为3级
    base_df['延期时长情况：0；≤1；1<X≤3；>3'] = base_df['节点审批延期时长(实际工作时长-节点审批时长）'].apply(
        lambda x: '0' if x == 0 else ('≤1' if x <= 1 else ('1<X≤3' if 1 < x <= 3 else '>3'))
    )

    # 返回处理结果
    return {
        'merged_data': base_df,
        'holiday_dates': holidays,
        'node_duration_map': node_duration_map
    }

if __name__ == '__main__':
    result = process_approval_data()
    print("数据处理完成")
    print("春节假期日期:", result['holiday_dates'])
    print("流程名称与时长对照:", result['node_duration_map'])
    print("处理后的数据的列名:")
    print(result['merged_data'].columns)
    merged_data = result['merged_data']
    merged_data.to_excel(r'E:\000000我的事项\2025-08\1131统计\审批结果.xlsx', index=False)
    print("处理后的数据已保存到审批结果1.xlsx")
    
    print("报告已生成并保存到审批结果.xlsx")



