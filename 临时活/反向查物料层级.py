import pandas as pd
from collections import defaultdict
from tqdm import tqdm

def reverse_material_hierarchy(input_file, output_file, target_items=None):
    # 读取Excel文件
    df = input_file.copy()
    
    # 构建物料编码与名称的映射
    material_name_map = {}
    for _, row in tqdm(df.iterrows(), desc="构建物料映射", total=len(df)):
        material_name_map[str(row['子项物料编码']).strip()] = str(row.get('子项物料描述', '')).strip()
        material_name_map[str(row['父项物料编码']).strip()] = str(row.get('父项物料描述', '')).strip()
    
    # 构建父子关系映射（修改版）
    child_parent_map = defaultdict(list)
    for _, row in tqdm(df.iterrows(), desc="构建父子关系", total=len(df)):
        child = str(row['子项物料编码']).strip()
        parent = str(row['父项物料编码']).strip()
        child_parent_map[child].append(parent)  # 保留所有父子关系
    
    print("完整父子关系映射示例:", child_parent_map['110200160373B'])
    
    # 结果数据结构
    result_data = []
    max_level = 0  # 记录最大层级深度
    
    # 处理每个目标子项
    for child in tqdm(target_items, desc="处理目标子项"):
        if child not in child_parent_map:
            print(f"子项 {child} 没有父项，直接作为顶级项处理")
            result_data.append({
                '子项物料': child,
                '子项物料名称': material_name_map.get(child, '')
            })
            continue
            
        print(f"\n处理子项: {child}")
        first_level_parents = child_parent_map.get(child, [])
        print(f"第一层父项({len(first_level_parents)}个): {', '.join(first_level_parents)}")
        
        # 使用广度优先搜索(BFS)遍历所有路径
        queue = [(child, [])]  # (当前物料, 路径)
        paths = []
        
        while queue:
            current_item, path = queue.pop(0)
            parents = child_parent_map.get(current_item, [])
            
            if not parents:
                paths.append(path)
                if len(path) > max_level:
                    max_level = len(path)
                continue
                
            for parent in parents:
                new_path = path + [parent]
                if len(new_path) == 1:  # 第一层父项处理
                    print(f"  子项 {child} -> 第一层父项 {parent}")
                queue.append((parent, new_path))
        
        # 将路径转换为结果行
        for path in paths:
            row = {
                '子项物料': child,
                '子项物料名称': material_name_map.get(child, '')
            }
            for level, parent in enumerate(path, 1):
                row[f'{level}级父'] = parent
                row[f'{level}级父名称'] = material_name_map.get(parent, '')
            result_data.append(row)
    
    # 统一所有行的列数
    columns = ['子项物料', '子项物料名称']
    for level in range(1, max_level + 1):
        columns.extend([f'{level}级父', f'{level}级父名称'])
    
    # 创建结果DataFrame
    result_df = pd.DataFrame(result_data)
    
    # 添加缺失的列
    for col in columns:
        if col not in result_df.columns:
            result_df[col] = None
    
    # 按指定列顺序输出
    result_df = result_df[columns]
    
    # 添加全量去重
    result_df = result_df.drop_duplicates()
    
    # 创建最高级父项对照表
    top_parent_map = []
    for child in target_items:
        if child not in child_parent_map:
            top_parent_map.append({
                '子项物料': child,
                '最高级父项': child,
                '子项物料名称': material_name_map.get(child, ''),
                '最高级父项名称': material_name_map.get(child, '')
            })
            continue
            
        # 使用BFS找出所有路径的最高级父项
        queue = [(child, [])]
        unique_top_parents = set()
        
        while queue:
            current_item, path = queue.pop(0)
            parents = child_parent_map.get(current_item, [])
            
            if not parents:  # 没有父项，说明是顶级
                if path:  # 如果有路径
                    unique_top_parents.add(path[-1])  # 取路径最后一个
                else:  # 没有路径，说明本身就是顶级
                    unique_top_parents.add(current_item)
            else:
                for parent in parents:
                    new_path = path + [parent]
                    queue.append((parent, new_path))
        
        # 记录所有唯一最高级父项
        for top_parent in unique_top_parents:
            top_parent_map.append({
                '子项物料': child,
                '最高级父项': top_parent,
                '子项物料名称': material_name_map.get(child, ''),
                '最高级父项名称': material_name_map.get(top_parent, '')
            })

    # 写入Excel前添加去重
    top_parent_df = pd.DataFrame(top_parent_map).drop_duplicates()

    # 写入Excel
    with pd.ExcelWriter(output_file) as writer:
        # 原有层级关系sheet
        result_df.to_excel(writer, sheet_name='层级关系', index=False)
        
        # 新增最高级父项对照sheet
        top_parent_df = pd.DataFrame(top_parent_map)
        top_parent_df.to_excel(writer, sheet_name='最高级父项', index=False)

# 使用示例
if __name__ == "__main__":
    # 凭借mbom数据并保留，父和子并去重
    df_mbom1 = pd.read_excel(fr"E:\000000我的事项\202606\模具报废\mbom-1000-1002-26-0115.XLSX")
    df_mbom2 = pd.read_excel(fr"E:\000000我的事项\202606\模具报废\mbom-1003-1005-26-0115.XLSX")
    df_mbom = pd.concat([df_mbom1, df_mbom2], axis=0)
    df_mbom = df_mbom[['物料编码', '组件']].drop_duplicates().reset_index(drop=True)
    df_mbom = df_mbom.rename(columns={'物料编码': '父项物料编码', '组件': '子项物料编码'})
    df_mbom[['父项物料编码', '子项物料编码']] = df_mbom[['父项物料编码', '子项物料编码']].astype(str)
    
    # 匹配进去物料名称
    df_pro_name1 = pd.read_excel(fr"E:\000000我的事项\202606\模具报废\物料-描述-冻结.XLSX")
    df_pro_name2 = pd.read_excel(fr"E:\000000我的事项\202606\模具报废\物料-描述-未冻结.XLSX")
    df_pro_name = pd.concat([df_pro_name1, df_pro_name2], axis=0)
    df_pro_name = df_pro_name[['物料编码', '物料描述']].drop_duplicates().reset_index(drop=True)
    df_pro_name[['物料编码', '物料描述']] = df_pro_name[['物料编码', '物料描述']].astype(str)
    
    df_mbom['父项物料描述'] = df_mbom['父项物料编码'].map(dict(zip(df_pro_name['物料编码'], df_pro_name['物料描述'])))
    df_mbom['子项物料描述'] = df_mbom['子项物料编码'].map(dict(zip(df_pro_name['物料编码'], df_pro_name['物料描述'])))
    input_excel = df_mbom
    print(input_excel.info())

    # input_excel = r"C:\Users\zhangbon\Desktop\模具报废-1231\所有bom的父子关系.xlsx"  # 替换为输入文件路径
    output_excel = r"E:\000000我的事项\202606\模具报废\电一模具.xlsx"  # 输出文件路径
    target_file = pd.read_excel(fr"E:\000000我的事项\202606\模具报废\电一模具报废钣金件物料号.xlsx")
    target_items = target_file['电一钣金物料号'].astype(str).str.strip().tolist()
    print(target_items)
    reverse_material_hierarchy(input_excel, output_excel, target_items=target_items)
