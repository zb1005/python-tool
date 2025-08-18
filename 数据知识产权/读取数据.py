import pandas as pd
import os

# 读取原始Workbook.xlsx
wb_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\Workbook-0815.xlsx")

# 目标文件夹路径（根据实际路径调整）
target_folder = r'C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\收入审核预测-更改版\省份'

# 遍历数据名称，查找对应文件并提取信息
result = {}
for 数据名称 in wb_df['数据名称'].tolist():
    file_name = f'{数据名称}.xlsx'
    found = False
    for dirpath, _, filenames in os.walk(target_folder):
        if file_name in filenames:
            file_path = os.path.join(dirpath, file_name)
            sheet = pd.read_excel(file_path)
            # print(sheet)
            print(sheet.iloc[16, -1])
            print('=====')
            应用场景 = sheet.iloc[1, -1] if len(sheet.iloc[1, -1]) > 10 else '无'
            算法 = sheet.iloc[16, -1] if len(sheet.iloc[16, -1]) > 10 else '无'
            result[数据名称] = {'应用场景': 应用场景, '算法': 算法, '文件路径': file_path}
            found = True
            break
    if not found:
        result[数据名称] = {'应用场景': '无', '算法': '无', '文件路径': '未找到'}

# 将结果转换为DataFrame并匹配到原始数据
result_df = pd.DataFrame.from_dict(result, orient='index').reset_index()
result_df.rename(columns={'index': '数据名称'}, inplace=True)

# 合并原始数据与结果
merged_df = pd.merge(wb_df, result_df, on='数据名称', how='left')

# 输出到新Excel文件（避免覆盖原文件）
merged_df.to_excel(r"C:\Users\zhangbon\Desktop\Workbook_算法匹配结果-0815.xlsx", index=False)
print('匹配结果已输出至Workbook_匹配结果.xlsx')