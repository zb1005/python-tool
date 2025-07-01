import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
import os

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

def process_approval_data(input_file_path):
    # 读取四个Excel文件
    base_df = pd.read_excel(input_file_path, sheet_name="主表", engine='openpyxl')
    staff_df = pd.read_excel(input_file_path, sheet_name="附1 最新在职人员及所属组织清单", engine='openpyxl')
    holiday_df = pd.read_excel(input_file_path, sheet_name="附2 方太春节假期", engine='openpyxl')
    special_node_df = pd.read_excel(input_file_path, sheet_name="附3特殊节点合理时长", engine='openpyxl')
    
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
    base_df['单个节点审批到达时间'] = pd.to_datetime(base_df['单个节点审批到达时间'])
    base_df['单个节点审批结束时间'] = pd.to_datetime(base_df['单个节点审批结束时间'])
    
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
    base_df['节点审批时效情况：≤1；1<X≤2；2<X≤3；>3'] = base_df['该节点审批工作时长_规整'].apply(
        lambda x: '≤1' if x <= 1 else ('1<X≤2' if 1 < x <= 2 else ('2<X≤3' if 2 < x <= 3 else '>3'))
    )

    # 7. 创建流程名称与建议时长的对照关系
    node_duration_map = dict(zip(
        special_node_df['流程名称'],
        special_node_df['建议合理时长（天）']
    ))
    
    # 8. 匹配节点合理审批时长，默认值为1
    base_df['节点审批时长'] = base_df['流程名称'].map(lambda x: node_duration_map.get(x, 1))

    # 9. 计算节点审批延期时长
    base_df['节点审批延期时长(实际工作时长-节点审批时长）'] = base_df['该节点审批工作时长'] - base_df['节点审批时长']
    base_df['节点审批延期时长(实际工作时长-节点审批时长）'] = base_df['节点审批延期时长(实际工作时长-节点审批时长）'].apply(lambda x: max(0, round(x, 2)))
    # 10. 判断节点审批延期时长情况，按照1，2，3天分为3级
    base_df['延期时长情况：≤1；1<X≤2；2<X≤3；>3'] = base_df['节点审批延期时长(实际工作时长-节点审批时长）'].apply(
        lambda x: '≤1' if x <= 1 else ('1<X≤2' if 1 < x <= 2 else ('2<X≤3' if 2 < x <= 3 else '>3'))
    )

    # 返回处理结果
    return {
        'merged_data': base_df,
        'holiday_dates': holidays,
        'node_duration_map': node_duration_map
    }

def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)

def run_processing():
    input_file = input_entry.get()
    if not input_file:
        messagebox.showerror("错误", "请选择输入文件")
        return
    
    try:
        result = process_approval_data(input_file)
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="审批结果.xlsx"
        )
        if not output_file:
            return  # 用户取消保存
        
        merged_data = result['merged_data']
        report1 = generate_report_1(merged_data)
        
        with pd.ExcelWriter(output_file) as writer:
            merged_data.to_excel(writer, sheet_name='原始数据', index=False)
            report1.to_excel(writer, sheet_name='按体系汇总', index=False)
        
        messagebox.showinfo("成功", f"报告已生成并保存到:\n{output_file}")
    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出错:\n{str(e)}")

if __name__ == '__main__':
    root = tk.Tk()
    root.title("审批计算工具")
    
    # 创建输入文件选择框架
    input_frame = tk.Frame(root, padx=10, pady=10)
    input_frame.pack(fill=tk.X)
    
    # 文件选择行
    file_select_frame = tk.Frame(input_frame)
    file_select_frame.pack(fill=tk.X)
    
    tk.Label(file_select_frame, text="输入文件:").pack(side=tk.LEFT)
    input_entry = tk.Entry(file_select_frame, width=50)
    input_entry.pack(side=tk.LEFT, padx=5)
    browse_btn = tk.Button(file_select_frame, text="浏览", command=select_input_file)
    browse_btn.pack(side=tk.LEFT)
    
    # Sheet要求说明
    sheet_requirements = "要求Excel包含以下sheet：主表、附1 最新在职人员及所属组织清单、附2 方太春节假期、附3特殊节点合理时长"
    tk.Label(input_frame, text=sheet_requirements, font=('Arial', 8), fg='gray').pack(anchor='w', pady=5)
    
    # 处理按钮
    process_btn = tk.Button(root, text="开始处理", command=run_processing, padx=20, pady=5)
    process_btn.pack(pady=10)
    
    root.mainloop()