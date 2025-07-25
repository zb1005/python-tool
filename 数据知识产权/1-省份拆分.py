import pandas as pd
import os

# 读取Excel文件
input_file = r"C:\Users\zhangbon\Desktop\数据产权\知识产权整理稿\选购场景分类数据\各省份案例数据.xlsx"
df = pd.read_excel(input_file)

# 获取所有唯一的省份值
provinces = df['省份'].unique()

# 获取输入文件所在目录
output_dir = os.path.dirname(input_file)

# 按省份拆分并保存为单独的Excel文件
for province in provinces:
    # 筛选当前省份的数据
    province_data = df[df['省份'] == province]
    # 构建输出文件名
    output_file = os.path.join(output_dir, f"{province}选购场景分类数据.xlsx")
    # 保存为Excel文件
    province_data.to_excel(output_file, index=False)

print(f"成功拆分{len(provinces)}个省份的数据，文件已保存至：{output_dir}")