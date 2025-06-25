"""
流程层级树形图绘制脚本

功能：读取Excel中的流程层级数据，绘制横向树形关系图

使用前请确保安装必要的库：
pip install pandas pyecharts openpyxl
"""

import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Tree

# 替换为实际的Excel文件路径
EXCEL_FILE_PATH = r"C:\Users\zhangbon\Desktop\美的-内销服务架构.xlsx"

def main():
    # 读取Excel数据
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
    except FileNotFoundError:
        print(f"错误：未找到文件 {EXCEL_FILE_PATH}，请检查文件路径是否正确。")
        return
    except Exception as e:
        print(f"读取Excel文件时出错：{e}")
        return

    # 检查是否包含必要的列
    required_columns = ["一级流程", "二级流程", "三级流程", "四级流程", "五级流程场景"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"错误：Excel文件缺少必要的列：{', '.join(missing_columns)}")
        return

    # 构建树形结构数据
    tree = {}
    for _, row in df.iterrows():
          current_node = tree
          for level in required_columns:
            value = row[level]
            # 跳过空值
            if pd.isna(value):
                break
            # 转换为字符串并去除空格
            value_str = str(value).strip()
            # 如果当前节点不存在该值，则创建新节点
            if value_str not in current_node:
                current_node[value_str] = {}
                
            # 移动到下一级节点
            current_node = current_node[value_str]

    # 将字典转换为pyecharts Tree所需的格式
    def dict_to_tree_data(node_dict):
        children = []
        for key, value in node_dict.items():
            if isinstance(value, dict) and value:
                children.append({"name": key, "children": dict_to_tree_data(value)})
            else:
                children.append({"name": key})
        return children

    tree_data = dict_to_tree_data(tree)

    # 如果树为空
    if not tree_data:
        print("错误：未从Excel数据中提取到有效层级关系。")
        return

    # 创建Tree图表，设置横向布局
    tree_chart = (
        Tree(init_opts=opts.InitOpts(width="1600px"))  # 增加图表宽度以扩大节点间距
        .add(
            series_name="流程层级",
            data=tree_data,
            orient="LR",  # LR表示横向布局（从左到右）
            layout={"verticalGap": 5000},
            label_opts=opts.LabelOpts(
                position="top",
                vertical_align="bottom",
                horizontal_align="center"
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="流程层级关系图"),
            tooltip_opts=opts.TooltipOpts(trigger="item", trigger_on="mousemove")
        )
    )

    # 渲染为HTML文件
    output_file = "流程层级关系图.html"
    tree_chart.render(output_file)
    print(f"树形图已成功生成：{output_file}")

if __name__ == "__main__":
    main()