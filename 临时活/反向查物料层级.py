import pandas as pd
from collections import defaultdict
from tqdm import tqdm

def reverse_material_hierarchy(input_file, output_file, target_items=None):
    # 读取Excel文件
    df = pd.read_excel(input_file)
    
    # 构建物料编码与名称的映射
    material_name_map = {}
    for _, row in tqdm(df.iterrows(), desc="构建物料映射", total=len(df)):
        material_name_map[str(row['子项物料编码']).strip()] = str(row.get('子项物料描述', '')).strip()
        material_name_map[str(row['父项物料编码']).strip()] = str(row.get('父项物料名称', '')).strip()
    
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
    input_excel = r"C:\Users\zhangbon\Desktop\3.XLSX"  # 替换为输入文件路径
    output_excel = r"C:\Users\zhangbon\Desktop\反向物料层级结果-燃气1.xlsx"  # 输出文件路径
    target_items = ['110200190100A', '110200190100B', '110200190099A', '110200190099B', '1102001606160', '110200060192A', '110200060192B', '1102001607120', '1102001607140', '1102001607220', '110200190555A', '110200190555C', '110200160565C', '110200190641A', '110200190641B', '1102001605660', '110200160568A', '110200190579A', '1102001606120', '1102001701730', '110200170185A', '110200160129C', '1102001608240', '1102001600420', '1102001601440', '1102001601450', '110200160145A', '110200160145B', '110200160201C', '110200100302B', '110200160373A', 
'110200160373B', '110200160313C', '110200160313D', '110200160313E', '1102001001330', '1102001401240', '110200180041A', '1102002001760', '110200130096A', '110200130096B', '110200130063A', '1102001301260', '110200130126A', '110200060138A', '110200060138C', '110200060138E', '110200060138F', '110200160143E', '110200160253A', '1102001803610', '1102001902130', '110200190213A', '110200190263C', '110200190263D', '110200190365B', '110200260264A', '1102003400920', '110200160143F', '110200160253B', '1102001608300', '110200260264B', '110200140044C', '110200140044D', '110200140044E', '110200140137A', '110200140137C', '1102001800500', '1102002000350', '110200150253E', '1104000103980', '110400020265A', '110400030439B', '110400031026A', '110400020266C', '1104000300330', '110400030439A', '1104000308920', '1104000310260', '110200200238A', '110200200238B', '110200200238C', '110200200238D', '110200200238E', '110200200238F', '1101001101220', '110200210084A', '1102001803310', '1102003401910', '110400030392A', '1104000100490', '110400010377B', '110400010377F', '110400010377H', '110400010427G', '110400010427H', '110400010427I', 
'110400010685D', '110400040081A', '110400040081C', '110400010427E', '110400010685C', '1104000400810', '1104000106860', '110400010378A', '110400010378B', '110400010378C', '110400030892A', '110400030972A', '1104000309720', '1104000304440', '1104000305750', '110400030575A', '110400030438A', '110400030438B', '1104000305700', '110400030570A', '110400030440A', '1104000305720', '110400030572A', '1104000305710', '110400030571A', '1104000300360', '110400060039B', '110400060037A', '110400060032A', '110400060032B', '110400060032C', '110400060032E', '110400060033A', '110400060033C', '110400060036D', '110400060038C', '1102002502500', '1102002600430', '110200260043B', '1102002606090', '1102002606100', '110200260040D', '110200260041A', '110200260041B', '110200260041C', '1102002600480', '110200260096A', '1102002600970', '110200290030B', '110200290030C', '110200290114A', '110200290033A', '110200290033B', '110200290033C', '110200290033D', '110200290116C', '110200290116D', '1102002602190', '1102001302020']


    reverse_material_hierarchy(input_excel, output_excel, target_items=target_items)
