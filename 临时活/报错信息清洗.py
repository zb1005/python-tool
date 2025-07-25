import re
import pandas as pd

def extract_bom_features(text):
    # 正则表达式匹配模式：
    pattern = r"\d{12}[A-Z0-9]"  # 匹配13位: 前12位数字 + 第13位数字/大写字母
    # 查找所有匹配项
    matches = re.findall(pattern, text)
    return matches
def extract_bom_features2(text):
    # 正则表达式匹配模式：
    # 匹配 QX 开头，结尾要是数字
    pattern = r"QX\d+"  # 匹配 QX 开头，结尾要是数字
    # 查找所有匹配项
    matches = re.findall(pattern, text)
    return matches
def extract_bom_features3(text):
    # 正则表达式匹配模式：
    # 匹配 QX 开头，结尾要是数字
    pattern = r"TFT\d+"  # 匹配 QX 开头，结尾要是数字
    # 查找所有匹配项
    matches = re.findall(pattern, text)
    return matches

# 示例用法
if __name__ == "__main__":
    df2 = pd.DataFrame()
    df = pd.read_excel(r"C:\Users\zhangbon\Desktop\export-MDM全量报错BOM的清单.XLSX",sheet_name="Sheet2")
    # 示例输入文本
    for index, row in df.iterrows():
        input_text = row["消息文本"]
        # 提取特征  
        results = extract_bom_features(input_text)
        results2 = extract_bom_features2(input_text)
        results3 = extract_bom_features3(input_text)
        print(results)
        print(results2)
        print(results3)
        for sub_material in results:
            print(f"物料：{sub_material}")
            df2 = pd.concat([df2, pd.DataFrame([{"物料":sub_material}])], ignore_index=True)
        for sub_material in results2:
            print(f"更改号：{sub_material}")
            df2 = pd.concat([df2, pd.DataFrame([{"更改号":sub_material}])], ignore_index=True)
        for sub_material in results3:
            print(f"更改号：{sub_material}")
            df2 = pd.concat([df2, pd.DataFrame([{"更改号":sub_material}])], ignore_index=True)

    df2.to_excel(r"C:\Users\zhangbon\Desktop\export-MDM全量报错BOM的清单提取.xlsx",index=False)

        
