import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd

class ApprovalAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("审批时效分析工具")
        
        # 文件选择框架
        self.file_frame = ttk.LabelFrame(root, text="文件选择")
        self.file_frame.pack(padx=10, pady=10, fill="x")
        
        # 4个文件选择入口
        self.file_paths = {}
        self.file_labels = ["主表", "在职人员清单", "假期表", "特殊节点表"]
        
        for i, label in enumerate(self.file_labels):
            frame = ttk.Frame(self.file_frame)
            frame.pack(fill="x", pady=2)
            
            ttk.Label(frame, text=f"{label}:").pack(side="left")
            ttk.Entry(frame, width=40).pack(side="left", padx=5)
            ttk.Button(frame, text="浏览", command=lambda idx=i: self.select_file(idx)).pack(side="left")
            
            self.file_paths[i] = {"label": label, "entry": frame.children["!entry"], "data": None}
        
        # Sheet预览框架
        self.sheet_frame = ttk.LabelFrame(root, text="Sheet预览")
        self.sheet_frame.pack(padx=10, pady=10, fill="x")
        
        self.sheet_combos = []
        for i in range(4):
            frame = ttk.Frame(self.sheet_frame)
            frame.pack(fill="x", pady=2)
            
            ttk.Label(frame, text=f"{self.file_labels[i]} Sheet:").pack(side="left")
            combo = ttk.Combobox(frame, state="readonly")
            combo.pack(side="left", padx=5, fill="x", expand=True)
            combo.bind("<<ComboboxSelected>>", lambda e, idx=i: self.load_columns(idx))
            self.sheet_combos.append(combo)
        
        # 列名预览框架
        self.column_frame = ttk.LabelFrame(root, text="列名预览")
        self.column_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.column_text = tk.Text(self.column_frame, height=10)
        self.column_text.pack(fill="both", expand=True)
        
        # 进度条
        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)
        
        # 输出路径选择
        self.output_frame = ttk.Frame(root)
        self.output_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(self.output_frame, text="输出路径:").pack(side="left")
        self.output_entry = ttk.Entry(self.output_frame, width=40)
        self.output_entry.pack(side="left", padx=5)
        ttk.Button(self.output_frame, text="浏览", command=self.select_output_path).pack(side="left")
        
        # 分析按钮
        ttk.Button(root, text="开始分析", command=self.analyze).pack(pady=10)
    
    def select_file(self, idx):
        filepath = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if filepath:
            self.file_paths[idx]["entry"].delete(0, tk.END)
            self.file_paths[idx]["entry"].insert(0, filepath)
            self.load_sheets(idx, filepath)
    
    def load_sheets(self, idx, filepath):
        try:
            self.progress["value"] = 25 * (idx + 1)
            self.root.update()
            
            xls = pd.ExcelFile(filepath)
            sheets = xls.sheet_names
            self.sheet_combos[idx]["values"] = sheets
            self.sheet_combos[idx].current(0)
            
            # 自动加载第一个sheet的列名
            self.load_columns(idx)
        except Exception as e:
            self.column_text.insert(tk.END, f"加载{self.file_labels[idx]}失败: {str(e)}\n")
    
    def load_columns(self, idx):
        filepath = self.file_paths[idx]["entry"].get()
        sheet = self.sheet_combos[idx].get()
        
        if filepath and sheet:
            try:
                df = pd.read_excel(filepath, sheet_name=sheet, nrows=1)
                self.column_text.insert(tk.END, f"{self.file_labels[idx]} - {sheet} 列名:\n")
                self.column_text.insert(tk.END, ", ".join(df.columns) + "\n\n")
            except Exception as e:
                self.column_text.insert(tk.END, f"加载列名失败: {str(e)}\n")
    
    def select_output_path(self):
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            title="选择输出文件路径"
        )
        if output_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_path)
    
    def analyze(self):
        output_path = self.output_entry.get()
        if not output_path:
            self.column_text.insert(tk.END, "请先选择输出文件路径！\n")
            return
            
        try:
            # 获取所有输入文件路径
            main_file = self.file_paths[0]["entry"].get()
            employee_file = self.file_paths[1]["entry"].get()
            leave_file = self.file_paths[2]["entry"].get()
            special_file = self.file_paths[3]["entry"].get()
            
            # 获取各文件的sheet名称
            main_sheet = self.sheet_combos[0].get()
            employee_sheet = self.sheet_combos[1].get()
            leave_sheet = self.sheet_combos[2].get()
            special_sheet = self.sheet_combos[3].get()
            
            # 加载所有数据
            self.progress["value"] = 25
            self.root.update()
            main_df = pd.read_excel(main_file, sheet_name=main_sheet)
            
            self.progress["value"] = 50
            self.root.update()
            employee_df = pd.read_excel(employee_file, sheet_name=employee_sheet)
            
            self.progress["value"] = 75
            self.root.update()
            leave_df = pd.read_excel(leave_file, sheet_name=leave_sheet)
            special_df = pd.read_excel(special_file, sheet_name=special_sheet)
            
            # 这里添加实际的审批计算逻辑
            # 示例: 计算审批时效
            result_df = self.calculate_approval_time(main_df, employee_df, leave_df, special_df)
            
            # 保存结果
            result_df.to_excel(output_path, index=False)
            
            self.progress["value"] = 100
            self.column_text.insert(tk.END, f"分析完成！结果已保存到: {output_path}\n")
        except Exception as e:
            self.column_text.insert(tk.END, f"分析过程中出错: {str(e)}\n")
            self.progress["value"] = 0

    def calculate_approval_time(self, main_df, employee_df, leave_df, special_df):
        # 1. 合并必要数据
        result_df = main_df.copy()
        
        # 2. 计算审批时效（示例逻辑，请根据实际需求调整）
        if '申请时间' in result_df.columns and '审批时间' in result_df.columns:
            result_df['审批时效(小时)'] = (pd.to_datetime(result_df['审批时间']) - 
                                     pd.to_datetime(result_df['申请时间'])).dt.total_seconds() / 3600
        
        # 3. 添加员工信息（示例）
        if '员工ID' in result_df.columns and '员工ID' in employee_df.columns:
            result_df = result_df.merge(employee_df, on='员工ID', how='left')
        
        # 4. 标记特殊节点（示例）
        if '节点ID' in result_df.columns and '节点ID' in special_df.columns:
            result_df = result_df.merge(special_df, on='节点ID', how='left')
        
        # 5. 处理假期影响（示例）
        if '申请时间' in result_df.columns and '假期开始' in leave_df.columns:
            # 这里可以添加假期影响的逻辑
            pass
        
        # 返回计算结果
        return result_df

if __name__ == "__main__":
    root = tk.Tk()
    app = ApprovalAnalysisApp(root)
    root.mainloop()