import datetime

def calculate_work_hours(start_dt, end_dt):
    # 计算总时间差（小时）
    total_hours = (end_dt - start_dt).total_seconds() / 3600
    
    # 初始化周末小时数
    weekend_hours = 0
    
    # 生成日期范围（按整天计算）
    current_date = start_dt.date()
    end_date = end_dt.date()
    
    while current_date <= end_date:
        # 判断是否为周末
        if current_date.weekday() >= 5:  # 5=周六，6=周日
            # 处理开始日期
            if current_date == start_dt.date():
                # 计算当天剩余的小时数
                next_day = current_date + datetime.timedelta(days=1)
                remaining = datetime.datetime.combine(next_day, datetime.time.min) - start_dt
                weekend_hours += remaining.total_seconds() / 3600
            # 处理结束日期
            elif current_date == end_dt.date():
                # 计算当天已过的小时数
                start_of_day = datetime.datetime.combine(current_date, datetime.time.min)
                passed = end_dt - start_of_day
                weekend_hours += passed.total_seconds() / 3600
            # 处理中间完整天数
            else:
                weekend_hours += 24
        current_date += datetime.timedelta(days=1)
    
    return total_hours - weekend_hours

def calculate_work_days(start_dt, end_dt):
    """
    计算两个日期之间的工作日天数差（排除周末）
    返回end_dt - start_dt的工作日天数差
    """
    # 确保输入是日期类型
    if isinstance(start_dt, datetime.datetime):
        start_dt = start_dt.date()
    if isinstance(end_dt, datetime.datetime):
        end_dt = end_dt.date()
    
    # 初始化工作日差计数器
    day_diff = 0
    current_date = start_dt
    
    # 计算日期差（包含end_dt当天）
    while current_date < end_dt:
        if current_date.weekday() < 5:  # 周一到周五
            day_diff += 1
        current_date += datetime.timedelta(days=1)
    
    # 如果是倒序日期（end_dt < start_dt）
    if end_dt < start_dt:
        current_date = end_dt
        while current_date < start_dt:
            if current_date.weekday() < 5:
                day_diff -= 1
            current_date += datetime.timedelta(days=1)
    
    return day_diff

# 使用示例
start_time = datetime.datetime(2023, 1, 6, 18, 0)  # 周五18:00
end_time = datetime.datetime(2023, 1, 9, 9, 0)    # 周一9:00
print(calculate_work_hours(start_time, end_time))  # 输出15.0
print(calculate_work_days(start_time, end_time))   # 输出1（周一）