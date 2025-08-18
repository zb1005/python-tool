import pandas as pd
from collections import defaultdict
from tqdm import tqdm


def forward_material_hierarchy(input_file, output_file, shaixuan_list):

    # 读取Excel文件
    df = pd.read_excel(input_file)
    df['父项物料编码'] = df['父项物料编码'].astype(str)
    df['子项物料编码'] = df['子项物料编码'].astype(str)
    
    # 构建物料编码与名称的映射
    material_name_map = {}
    for _, row in tqdm(df.iterrows(), desc="构建物料映射", total=len(df)):
        material_name_map[str(row['父项物料编码']).strip()] = str(row.get('父项物料名称', '')).strip()
        material_name_map[str(row['子项物料编码']).strip()] = str(row.get('子项物料描述', '')).strip()
    
    # 构建父到子的关系映射
    parent_child_map = defaultdict(list)
    all_children = set()
    for _, row in tqdm(df.iterrows(), desc="构建父子关系", total=len(df)):
        parent = str(row['父项物料编码']).strip()
        child = str(row['子项物料编码']).strip()
        parent_child_map[parent].append(child)
        all_children.add(child)
    
    # 找出所有顶级父项（没有父项的物料）
    all_parents = set(parent_child_map.keys())
    top_level_parents = all_parents - all_children
    print(f"找到 {len(top_level_parents)} 个顶级父项")
    #引入了一个筛选，只对这部分进行分析
    top_level_parents = shaixuan_list  
    print(len(top_level_parents))

    # 结果数据结构
    result_data = []
    # 处理每个顶级父项
    for parent in tqdm(top_level_parents, desc="处理顶级父项"):
        # 使用广度优先搜索(BFS)向下展开所有子项
        queue = [(parent, [parent])]  # (当前物料, 完整路径)
        
        while queue:
            current_item, path = queue.pop(0)
            children = parent_child_map.get(current_item, [])

            # 将子项加入队列
            for child in children:
                row = {}
                row['父项物料编码'] = current_item
                row['子项物料编码'] = child
                result_data.append(row.copy())
                new_path = [child]
                queue.append((child, new_path))
    
    # 创建结果DataFrame
    result_df = pd.DataFrame(result_data)
     
    # 添加全量去重
    result_df = result_df.drop_duplicates()
    
    # 创建最低级子项对照表
    bottom_child_map = []
    for parent in top_level_parents:
        # 使用BFS找出所有路径的最低级子项
        queue = [(parent, [])]
        unique_bottom_children = set()
        
        while queue:
            current_item, path = queue.pop(0)
            children = parent_child_map.get(current_item, [])
            
            if not children:  # 没有子项，说明是最低级
                if path:  # 如果有路径
                    unique_bottom_children.add(path[-1])  # 取路径最后一个
                else:  # 没有路径，说明本身就是最低级
                    unique_bottom_children.add(current_item)
            else:
                for child in children:
                    new_path = path + [child]
                    queue.append((child, new_path))
    
        # 记录所有唯一最低级子项
        for bottom_child in unique_bottom_children:
            bottom_child_map.append({
                '父项物料': parent,
                '最低级子项': bottom_child,
                '父项物料名称': material_name_map.get(parent, ''),
                '最低级子项名称': material_name_map.get(bottom_child, '')
            })

    # 写入Excel前添加去重
    bottom_child_df = pd.DataFrame(bottom_child_map).drop_duplicates()

    # 写入Excel
    with pd.ExcelWriter(output_file) as writer:
        # 层级关系sheet
        result_df.to_excel(writer, sheet_name='层级关系', index=False)
        # 最低级子项对照sheet
        bottom_child_df.to_excel(writer, sheet_name='最低级子项', index=False)

# 使用示例
if __name__ == "__main__":
    input_excel = r"C:\Users\zhangbon\Desktop\export-全量.XLSX"  # 替换为输入文件路径
    output_excel = r"C:\Users\zhangbon\Desktop\正向物料层级结果.xlsx"  # 输出文件路径
    
    筛选_excel = pd.read_excel(r'C:\Users\zhangbon\Desktop\工程产品清单0721.xlsx',sheet_name='Sheet1')
    筛选_excel['产品编码'] = 筛选_excel['产品编码'].astype(str)
    shaixuan_list = 筛选_excel['产品编码'].tolist()

    forward_material_hierarchy(input_excel, output_excel, shaixuan_list)
