import pandas as pd
# 读取Excel文件（请替换实际路径）
df = pd.read_excel(r'C:\Users\zhangbon\Desktop\油烟机和燃气验证数据用.xlsx',sheet_name='燃气工厂')

# 转换为扁平化一维列表并将所有元素转为字符串，同时过滤特殊字符和换行符
data_list = [str(item).replace('\n', '').replace('\r', '').strip() for item in df.values.flatten().tolist()]

print(data_list)
# print(len(data_list[:20]))

