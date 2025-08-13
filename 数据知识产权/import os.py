import os
import pandas as pd
import re
from pathlib import Path

def batch_rename_columns():
    """
    批量重命名Excel文件的列名
    使用正则表达式精确匹配文件名
    """
    
    # 目标文件夹路径
    base_dir = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\省份"
    
    # 定义列名映射规则
    column_mapping = {
        "2024年7月审核金额(万元)": "预测月前6个月审核金额(万元)",
        "2024年8月审核金额(万元)": "预测月前5个月审核金额(万元)",
        "2024年9月审核金额(万元)": "预测月前4个月审核金额(万元)",
        "2024年10月审核金额(万元)": "预测月前3个月审核金额(万元)",
        "2024年11月审核金额(万元)": "预测月前2个月审核金额(万元)",
        "2024年12月审核金额(万元)": "预测月前1个月审核金额(万元)",
        "预测完成日期": "预测日期",
        "2025年1月预测金额(万元)": "本月预测金额(万元)",
        "本月预测金额(万元)":'本月预测销售金额(万元)'
    }
        
    # 定义正则表达式模式
    patterns = [
        r".*门店洗碗机.*\.xlsx$",
        r".*门店油烟机.*\.xlsx$",
        r".*门店灶具.*\.xlsx$"
    ]
    
    # 记录处理结果
    processed_files = []
    error_files = []
    
    # 检查基础目录是否存在
    if not os.path.exists(base_dir):
        print(f"错误：目录 {base_dir} 不存在")
        return
    
    # 遍历省份文件夹下的所有子文件夹
    for province_folder in os.listdir(base_dir):
        province_path = os.path.join(base_dir, province_folder)
        
        # 确保是文件夹
        if not os.path.isdir(province_path):
            continue
            
        print(f"\n正在处理省份：{province_folder}")
        
        # 查找符合条件的Excel文件
        excel_files = []
        
        # 遍历文件夹中的所有文件
        for filename in os.listdir(province_path):
            if filename.endswith('.xlsx') and not filename.startswith('~$'):
                # 检查是否匹配任一正则表达式
                for pattern in patterns:
                    if re.match(pattern, filename, re.IGNORECASE):
                        full_path = os.path.join(province_path, filename)
                        excel_files.append(full_path)
                        break
        
        if not excel_files:
            print(f"  省份 {province_folder} 中没有找到符合条件的Excel文件")
            continue
            
        print(f"  找到 {len(excel_files)} 个Excel文件")
        
        # 处理每个Excel文件
        for file_path in excel_files:
            try:
                # 读取Excel文件
                df = pd.read_excel(file_path, engine='openpyxl')
                
                # 获取原始列名
                original_columns = df.columns.tolist()
                
                # 应用列名映射
                df.rename(columns=column_mapping, inplace=True)
                
                # 检查是否有列名被修改
                new_columns = df.columns.tolist()
                modified = False
                
                for old_name, new_name in column_mapping.items():
                    if old_name in original_columns:
                        modified = True
                        print(f"    重命名列：'{old_name}' -> '{new_name}'")
                
                if modified:
                    # 保存修改后的文件（覆盖原文件）
                    df.to_excel(file_path, index=False, engine='openpyxl')
                    processed_files.append(file_path)
                    print(f"  ✓ 已处理：{os.path.basename(file_path)}")
                else:
                    print(f"  - 跳过：{os.path.basename(file_path)}（未找到需要重命名的列）")
                    
            except Exception as e:
                error_files.append((file_path, str(e)))
                print(f"  ✗ 处理失败：{os.path.basename(file_path)} - {str(e)}")
    
    # 输出处理结果总结
    print("\n" + "="*50)
    print("处理结果总结：")
    print(f"成功处理文件数：{len(processed_files)}")
    print(f"处理失败文件数：{len(error_files)}")
    
    if processed_files:
        print("\n成功处理的文件：")
        for file in processed_files:
            print(f"  - {os.path.basename(file)}")
    
    if error_files:
        print("\n处理失败的文件：")
        for file, error in error_files:
            print(f"  - {os.path.basename(file)}: {error}")

if __name__ == "__main__":
    print("开始批量重命名Excel列名...")
    print("规则：将2024年7-12月审核金额改为预测月前6-1个月审核金额")
    print("文件匹配规则：使用正则表达式匹配包含'门店洗碗机'、'门店油烟机'、'门店灶具'的.xlsx文件")
    batch_rename_columns()
    print("\n处理完成！")