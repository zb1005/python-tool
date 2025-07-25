import os
import pandas as pd
from pathlib import Path

# 目标省份文件夹路径（根据用户实际路径调整）
base_dir = Path(r"c:\Users\zhangbon\Desktop\python-tool\省份")

# 定义列名映射（前推n月 -> 具体月份）
column_mapping = {
    "前推6个月审核金额": "2024年7月审核金额(万元)",
    "前推5个月审核金额": "2024年8月审核金额(万元)",
    "前推4个月审核金额": "2024年9月审核金额(万元)",
    "前推3个月审核金额": "2024年10月审核金额(万元)",
    "前推2个月审核金额": "2024年11月审核金额(万元)",
    "前推1个月审核金额": "2024年12月审核金额(万元)",
    "前推6个月审核金额(万元)": "2024年7月审核金额(万元)",
    "前推5个月审核金额(万元)": "2024年8月审核金额(万元)",
    "前推4个月审核金额(万元)": "2024年9月审核金额(万元)",
    "前推3个月审核金额(万元)": "2024年10月审核金额(万元)",
    "前推2个月审核金额(万元)": "2024年11月审核金额(万元)",
    "前推1个月审核金额(万元)": "2024年12月审核金额(万元)",
    "预测金额(万元)": "2025年1月预测金额(万元)"
}

# 定义产品类型及分配比例（油烟机:灶具:洗碗机 = 15:10:4）
products = [
    {"name": "油烟机", "ratio": 15},
    {"name": "灶具", "ratio": 10},
    {"name": "洗碗机", "ratio": 4}
]

def process_province_excel(province_dir):
    # 查找目标Excel文件（过滤临时文件~$开头）
    for file in province_dir.glob("*门店销售金额预测数据.xlsx"):
        if file.name.startswith("~$"):  # 跳过Excel临时文件
            continue
        # 读取原始数据
        df = pd.read_excel(file, engine="openpyxl")
        # 重命名列
        df.rename(columns=column_mapping, inplace=True)
        # 获取「2024年7月审核金额(万元)」的列索引
        target_col = f"2024年7月审核金额(万元)"
        target_index = df.columns.get_loc(target_col)
        # 计算总比例
        total_ratio = sum(p["ratio"] for p in products)
        # 为每个产品生成新数据
        for product in products:
            new_df = df.copy()
            # 插入「销售产品产品类」列至目标列前
            new_df.insert(target_index, "销售产品产品类", product["name"])
            # 按比例分配数值（所有数值列均按比例调整）
            value_columns = [col for col in new_df.columns if "金额(万元)" in col]
            new_df[value_columns] = new_df[value_columns] * (product["ratio"] / total_ratio)
            # 新增：数值四舍五入到两位小数
            new_df[value_columns] = new_df[value_columns].round(2)

            # 新增：文本内容替换（遍历所有字符串列）
            for col in new_df.columns:
                if new_df[col].dtype == 'object':  # 仅处理字符串列
                    # 替换规则1：办事处-门店描述
                    new_df[col] = new_df[col].str.replace(
                        "各个办事处下的各个门店过去半年度的审核订单金额数据",
                        f"各个办事处下的各个门店预测月前六个月的{product['name']}产品的审核订单金额数据"
                    )
                    # 替换规则2：数据收集描述
                    new_df[col] = new_df[col].str.replace(
                        "数据收集：收集过去半年录入订单的历史数据，以门店为颗粒度，涵盖各门店的订单金额、时间维度及地区属性等信息。",
                        f"数据收集：收集预测月前六个月录入的{product['name']}订单的历史数据，以门店为颗粒度，涵盖各门店的订单金额、时间维度及地区属性等信息。"
                    )
                    #替换规则3：数据名称替换
                    new_df[col] = new_df[col].str.replace(
                        "门店销售金额预测数据",
                        f"门店{product['name']}销售金额预测数据"
                    )

            # 生成新文件名（已包含产品类型，符合规则3）
            new_filename = file.name.replace("门店销售金额预测数据", f"门店{product['name']}销售金额预测数据")
            new_path = province_dir / new_filename
            # 保存新文件
            new_df.to_excel(new_path, index=False, engine="openpyxl")
        print(f"处理完成：{file.name}")

# 遍历所有省份子文件夹
for province_dir in base_dir.iterdir():
    if province_dir.is_dir():
        process_province_excel(province_dir)