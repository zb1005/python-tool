import pandas as pd
from datetime import datetime, timedelta

def is_workday(date, holidays):
    # 判断是否为工作日（非周末且非节假日）
    return date.weekday() < 5 and date not in holidays

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

def generate_report_1(merged_data):
    """生成第一个汇总报告：按审批人所在体系分组统计"""
    # 确保分组列存在
    if '审批人所在体系' not in merged_data.columns:
        raise ValueError("数据中缺少'审批人所在体系'列")
    
    # 确保所有需要的列都存在
    required_columns = ['流程名称', '节点审批时效情况（≤1；＞1）', 
                       '节点审批时效是否大于3天', '该节点审批自然时长（单位：天）',
                       '该节点审批工作时长（单位：天）——剔除节假日及周末，按24小时计算']
    missing_cols = [col for col in required_columns if col not in merged_data.columns]
    if missing_cols:
        raise ValueError(f"数据中缺少必要的列: {missing_cols}")
    
    report = merged_data.groupby('审批人所在体系').agg(
        **{
            'A-审批总次数': ('流程名称', 'count'),
            'B-单个节点审批时长≤1天的节点数': ('节点审批时效情况（≤1；＞1）', lambda x: (x == '<=1').sum()),
            'D-单个节点审批时长＞1天的节点数': ('节点审批时效情况（≤1；＞1）', lambda x: (x == '>1').sum()),
            'E-其中：单个节点审批时长>3天的节点数': ('节点审批时效是否大于3天', lambda x: (x == 'Y').sum()),
            'G-单个节点平均审批时长（自然日）': ('该节点审批自然时长（单位：天）', 'mean'),
            'H-单个节点平均审批时长（工作日）': ('该节点审批工作时长（单位：天）——剔除节假日及周末，按24小时计算', 'mean')
        }
    )
    
    # 计算比例列
    report['C-单个节点审批时长≤1天的节点比例'] = report['B-单个节点审批时长≤1天的节点数'] / report['A-审批总次数']
    report['F-其中：单个节点审批时长>3天的节点比例'] = report['E-其中：单个节点审批时长>3天的节点数'] / report['A-审批总次数']
    
    # 重置索引，将'审批人所在体系'变为普通列
    report = report.reset_index()
    
    # 确保列顺序正确
    report = report[[
        '审批人所在体系',
        'A-审批总次数',
        'B-单个节点审批时长≤1天的节点数',
        'C-单个节点审批时长≤1天的节点比例',
        'D-单个节点审批时长＞1天的节点数',
        'E-其中：单个节点审批时长>3天的节点数',
        'F-其中：单个节点审批时长>3天的节点比例',
        'G-单个节点平均审批时长（自然日）',
        'H-单个节点平均审批时长（工作日）'
    ]]
    
    # 保留2位小数
    report = report.round(2)
    
    return report

def process_approval_data():
    # 读取四个Excel文件
    base_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\临时活\1131统计\Q1审批时效数据分析打样-0513.xlsx",sheet_name="主表", engine='openpyxl')
    staff_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\临时活\1131统计\Q1审批时效数据分析打样-0513.xlsx",sheet_name="附1 最新在职人员及所属组织清单", engine='openpyxl')
    holiday_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\临时活\1131统计\Q1审批时效数据分析打样-0513.xlsx",sheet_name="附2 方太春节假期", engine='openpyxl')
    special_node_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\临时活\1131统计\Q1审批时效数据分析打样-0513.xlsx",sheet_name="附3特殊节点合理时长", engine='openpyxl')
    
    # 1. 先建立人员信息的映射字典
    staff_mapping = staff_df.set_index('员工工号').to_dict('index')
    print(staff_mapping)
    
    # 2. 将人员信息映射到基础表
    base_df['审批人姓名'] = base_df['审批人工号'].map(lambda x: staff_mapping.get(x, {}).get('姓名'))
    base_df['审批人所在体系'] = base_df['审批人工号'].map(lambda x: staff_mapping.get(x, {}).get('审批人所在体系'))
    base_df['审批人所在一级组织'] = base_df['审批人工号'].map(lambda x: staff_mapping.get(x, {}).get('一级组织名称'))
    
    # 3. 处理假期日期
    holidays = [datetime.strptime(str(date).strip(), '%Y-%m-%d %H:%M:%S').date() for date in holiday_df['方太假期']]
    
    # 4. 计算自然时长和工作时长
    base_df['单个节点审批到达时间（格式如：2024-09-26 15:20:59，精确到秒）'] = pd.to_datetime(base_df['单个节点审批到达时间（格式如：2024-09-26 15:20:59，精确到秒）'])
    base_df['单个节点审批结束时间（格式如：2024-09-26 15:20:59，精确到秒）'] = pd.to_datetime(base_df['单个节点审批结束时间（格式如：2024-09-26 15:20:59，精确到秒）'])
    
    base_df['该节点审批自然时长（单位：天）'] = (base_df['单个节点审批结束时间（格式如：2024-09-26 15:20:59，精确到秒）'] - base_df['单个节点审批到达时间（格式如：2024-09-26 15:20:59，精确到秒）']).dt.total_seconds() / (24 * 3600)
    
    base_df['该节点审批工作时长（单位：天）——剔除节假日及周末，按24小时计算'] = base_df.apply(
        lambda row: calculate_work_duration(
            row['单个节点审批到达时间（格式如：2024-09-26 15:20:59，精确到秒）'], 
            row['单个节点审批结束时间（格式如：2024-09-26 15:20:59，精确到秒）'], 
            holidays
        ), 
        axis=1
    )
    
    # 5. 规整工作时长（保留2位小数），小于0时设为0
    base_df['该节点审批工作时长_规整'] = base_df['该节点审批工作时长（单位：天）——剔除节假日及周末，按24小时计算'].apply(lambda x: max(0, round(x, 2)))
    
    # 6. 创建流程名称与建议时长的对照关系
    node_duration_map = dict(zip(
        special_node_df['流程名称'],
        special_node_df['建议合理时长（天）']
    ))
    
    # 7. 匹配节点合理审批时长，默认值为1
    base_df['节点合理审批时长（天）'] = base_df['流程名称'].map(lambda x: node_duration_map.get(x, 1))

    # 8. 计算节点审批延期时长
    base_df['节点审批延期时长（天）'] = base_df['该节点审批工作时长（单位：天）——剔除节假日及周末，按24小时计算'] - base_df['节点合理审批时长（天）']
    base_df['节点审批延期时长（天）'] = base_df['节点审批延期时长（天）'].apply(lambda x: max(0, round(x, 2)))
    # 9. 判断节点审批时效情况，如果小于等于0则显示<=1，否则显示>1
    base_df['节点审批时效情况（≤1；＞1）'] = base_df['节点审批延期时长（天）'].apply(lambda x: '<=1' if x <= 0 else '>1')

    # 10. 判断节点审批时效是否大于3天
    base_df['节点审批时效是否大于3天'] = base_df['该节点审批工作时长（单位：天）——剔除节假日及周末，按24小时计算'].apply(lambda x: 'Y' if x > 3 else 'N')
    
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
    merged_data.to_excel(r'C:\Users\zhangbon\Desktop\审批结果1.xlsx', index=False)
    print("处理后的数据已保存到审批结果1.xlsx")
    # 生成报告1
    report1 = generate_report_1(result['merged_data'])
    
    # 将结果保存到Excel，包含多个sheet
    with pd.ExcelWriter(r'C:\Users\zhangbon\Desktop\审批结果.xlsx') as writer:
        result['merged_data'].to_excel(writer, sheet_name='原始数据', index=False)
        report1.to_excel(writer, sheet_name='按体系汇总', index=False)
    
    print("报告已生成并保存到审批结果.xlsx")



