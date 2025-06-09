import os
import pandas as pd

def batch_rename(folder_path, rename_rules):
    """
    批量重命名文件夹中的文件（精确前缀匹配版）
    :param folder_path: 文件夹绝对路径
    :param rename_rules: 包含替换规则的DataFrame
    """
    for filename in os.listdir(folder_path):
        for _, rule in rename_rules.iterrows():
            target_text = rule['文件名称']
            new_part = rule['文件编码']
            
            if target_text in filename:
                parts = filename.split(target_text, 1)
                
                # 精确匹配前缀并替换
                if parts[0].strip() != "":
                    # 直接使用新前缀替换整个旧前缀
                    new_name = f"{new_part}-{target_text}{parts[1]}"
                    
                    src = os.path.join(folder_path, filename)
                    dst = os.path.join(folder_path, new_name)
                    
                    # 执行重命名并更新当前文件名引用
                    os.rename(src, dst)
                    print(f"Renamed: {filename} -> {new_name}")
                    filename = new_name  # 更新引用防止多次处理

if __name__ == "__main__":
    # 从Excel读取规则（新格式）
    rule_file = r"C:\Users\zhangbon\Desktop\编码\待更新编码清单1.xlsx"
    rename_rules = pd.read_excel(rule_file)
    
    # 指定目标文件夹
    folder_path = r"C:\Users\zhangbon\Desktop\编码\主数据管理制度\附录C《主数据运维与管控流程体系》"
    
    batch_rename(folder_path, rename_rules)