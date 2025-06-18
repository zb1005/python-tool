import pandas as pd
import numpy as np
from scipy import stats
import openpyxl
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore")

def calculate_work_duration(start_time, end_time, holidays):
    # 确保输入是datetime类型
    def parse_time(time_str):
        if isinstance(time_str, str):
            try:
                return datetime.strptime(time_str, '%Y/%m/%d %H:%M:%S')
            except ValueError:
                try:
                    return datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    try:
                        return datetime.strptime(time_str, '%Y/%m/%d')
                    except ValueError:
                        return datetime.strptime(time_str, '%Y-%m-%d')
        return time_str
    
    start_time = parse_time(start_time)
    end_time = parse_time(end_time)
    
    current = start_time
    work_duration = timedelta()
    
    while current < end_time:
        next_day = (current + timedelta(days=1)).replace(hour=0, minute=0, second=0)
        if current.date() not in holidays:
            work_duration += min(next_day, end_time) - current
        current = next_day
    
    return work_duration.total_seconds() / (24 * 3600)  # 转换为天数

def analyze_approval_duration(df, group_field1, group_field2):
    """
    对审批时长进行分组统计和置信度分析
    :param df: 包含审批数据的DataFrame
    :param group_field1: 第一个分组字段名
    :param group_field2: 第二个分组字段名
    :return: 包含统计结果的DataFrame
    """
    # 分组计算计数、均值、标准差和分位数
    grouped = df.groupby([group_field1, group_field2])['该节点审批工作时长'].agg(
        ['count', 'mean', 'std', 
         lambda x: x.quantile(0.95), 
         lambda x: x.quantile(0.90),
         lambda x: x.quantile(0.80)]
    ).reset_index()
    grouped = grouped.rename(columns={
        '<lambda_0>': '95%分位',
        '<lambda_1>': '90%分位',
        '<lambda_2>': '80%分位'
    })
    
    # 计算60%、80%、90%和95%的置信区间
    # def calc_ci(row, confidence=0.95):
    #     if row['count'] <= 1 or pd.isna(row['std']):
    #         return (np.nan, np.nan)
    #     return stats.t.interval(
    #         confidence, 
    #         df=row['count']-1,
    #         loc=row['mean'],
    #         scale=row['std']/np.sqrt(row['count'])
    #     )
    
    # grouped['60%_CI'] = grouped.apply(lambda x: calc_ci(x, 0.60), axis=1)
    # grouped['80%_CI'] = grouped.apply(lambda x: calc_ci(x, 0.80), axis=1)
    # grouped['90%_CI'] = grouped.apply(lambda x: calc_ci(x, 0.90), axis=1)
    # grouped['95%_CI'] = grouped.apply(calc_ci, axis=1)
    
    # # 拆分置信区间为上下限
    # grouped['60%_CI_lower'] = grouped['60%_CI'].apply(lambda x: x[0])
    # grouped['60%_CI_upper'] = grouped['60%_CI'].apply(lambda x: x[1])
    # grouped['80%_CI_lower'] = grouped['80%_CI'].apply(lambda x: x[0])
    # grouped['80%_CI_upper'] = grouped['80%_CI'].apply(lambda x: x[1])
    # grouped['90%_CI_lower'] = grouped['90%_CI'].apply(lambda x: x[0])
    # grouped['90%_CI_upper'] = grouped['90%_CI'].apply(lambda x: x[1])
    # grouped['95%_CI_lower'] = grouped['95%_CI'].apply(lambda x: x[0])
    # grouped['95%_CI_upper'] = grouped['95%_CI'].apply(lambda x: x[1])
    
    # # 删除原始CI列
    # grouped.drop(['60%_CI', '80%_CI', '90%_CI', '95%_CI'], axis=1, inplace=True)
    
    return grouped

# 使用示例
# df = pd.read_excel("您的数据文件.xlsx")
# result = analyze_approval_duration(df, "字段1", "字段2")
# result.to_excel("分析结果.xlsx", index=False)

holiday_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\临时活\1131统计\计算节点合理时长\2025-方太非工作日清单.xlsx", engine='openpyxl')
holidays = [datetime.strptime(str(date).strip(), '%Y-%m-%d %H:%M:%S').date() for date in holiday_df['日期']]

in_path = r"C:\Users\zhangbon\Desktop\临时活\1131统计\计算节点合理时长\24年数据分析-0603\24年数据分析-0603\Q1-五个一和PBC审批流程(1).xlsx"
out_path = in_path.replace(".xlsx", "-计算时长.xlsx")
group_path = in_path.replace(".xlsx", "-结果分析稿.xlsx")


df = pd.read_excel(in_path, engine='openpyxl')


df['该节点审批工作时长'] = df.apply(
    lambda row: calculate_work_duration(
        row['单个节点审批到达时间'], 
        row['单个节点审批结束时间'], 
        holidays
    ), 
    axis=1
)

df.to_excel(out_path, index=False)
print(df.columns)
grouper = analyze_approval_duration(df, '流程名称', '审批节点名称')
grouper.to_excel(group_path, index=False)
print("分析完成")



# # 绘图
# import matplotlib.pyplot as plt
# import seaborn as sns
# #在绘图代码前添加中文字体设置
# plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置中文字体为黑体
# plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
# # 筛选特定流程和节点的数据
# filtered_data = df[(df['流程名称'] == 'PBC审批流程') & 
#                   (df['审批节点名称'] == '直接上级')]['该节点审批工作时长']

# # 绘制正态分布图
# plt.figure(figsize=(10, 6))
# sns.histplot(filtered_data, kde=True, stat='density', linewidth=0)
# sns.kdeplot(filtered_data, color='red', linewidth=2)

# # 添加图表标题和标签
# plt.title('PBC审批时长正态分布图', fontsize=15)
# plt.xlabel('审批时长(天)', fontsize=12)
# plt.ylabel('密度', fontsize=12)

# # 显示图表
# plt.grid(True)
# plt.tight_layout()
# plt.savefig(r"C:\Users\zhangbon\Desktop\临时活\1131统计\计算节点合理时长\PBC直接上级审批.png")
# plt.show()
