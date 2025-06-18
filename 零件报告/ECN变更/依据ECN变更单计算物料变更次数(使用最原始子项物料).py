import pandas as pd
import numpy as np
from collections import defaultdict

def count_material_changes(file_path,sheet_name='Sheet1'):
    print("============开始============")
    # 读取Excel文件
    df = pd.read_excel(file_path,sheet_name=sheet_name)
    print("============读取完成============")
    print("df:",df)
    print("df.columns:",df.columns,len(df))
    print("============开始处理============")
    # 剔除研发ECN字段为空或长度不足的行
    df = df[df['研发ECN'].apply(lambda x: len(str(x).strip()) >= 5)]
    df = df.reset_index(drop=True)  # 重置索引并丢弃旧索引
    print("已剔除研发ECN无效的行，剩余行数:", len(df))
    # 删除或修正下面这行错误的代码
    # df = df.reindex  # 错误的写法
    df = df.reindex()  # 正确的写法，但通常不需要这行
    
    # 初始化变更次数统计字典
    change_counts = defaultdict(int)
    # 物料变更链映射 {新物料: 旧物料}
    material_chains = {}
    
    # 遍历每一行数据
    for index, row in df.iterrows():
        # 获取当前行的物料编码
        current_material = str(row['子项物料']).strip()
        old_material = current_material  # 初始化old_material
        
        # 判断是否是变更后的物料（根据更改号自的长度）
        if pd.notna(row.get('更改号自')) and len(str(row['更改号自'])) >= 5:
            # 查找对应的旧物料行（同时匹配更改号至和父项物料）
            old_material_row = df[
                (df['更改号至'] == row['更改号自']) & 
                (df['父项物料编码'] == row['父项物料编码']) &
                (df['子项物料'] != current_material)  # 确保不是同一个物料
            ]
            
            if not old_material_row.empty:
                old_material = str(old_material_row.iloc[0]['子项物料']).strip()
                # 确保新旧物料不同
                if old_material != current_material:
                    material_chains[current_material] = old_material
                    print(f"建立变更关系: {old_material} -> {current_material}")
            else:
                print(f"警告: 找不到匹配的旧物料行 for {current_material}")
            print("current_material:",current_material)
            print("old_material:",old_material)
        print("index:",index)

    # 计算每个物料的完整变更链长度
    final_chains = {}
    # 新增：记录每个物料的最原始版本
    original_materials = {}
    
    # 先找出所有链条的起点（没有作为value出现的物料）
    chain_starts = set(material_chains.keys()) - set(material_chains.values())
    
    # 对每个链条起点计算整个链条长度
    for start in chain_starts:
        chain_length = 0
        current = start
        chain_materials = [current]
        
        # 向下遍历整个链条
        while current in material_chains:
            current = material_chains[current]
            chain_materials.append(current)
            chain_length += 1
        
        # 为链条中所有物料记录总变更次数
        for material in chain_materials:
            final_chains[material] = chain_length
            # 记录最原始物料（链条的最后一个物料）
            original_materials[material] = chain_materials[-1]

    # 新增：构建物料编码与名称的映射表
    material_name_map = {}
    for _, row in df.iterrows():
        material_code = str(row['子项物料']).strip()
        if material_code not in material_name_map:
            material_name_map[material_code] = row.get('子项物料名称', '')
        
        parent_code = str(row['父项物料编码']).strip()
        if parent_code not in material_name_map:
            material_name_map[parent_code] = row.get('父项物料名称', '')

    # 将变更次数映射回原始DataFrame
    df['变更次数'] = df['子项物料'].map(lambda x: final_chains.get(str(x).strip(), 0))
    # 新增最原始子项物料列
    df['最原始子项物料'] = df['子项物料'].map(lambda x: original_materials.get(str(x).strip(), str(x).strip()))
    
    # 新增是否在ADCP前列
    df['是否在ADCP前'] = df['研发ECN'].apply(lambda x: 'ADCP前' if str(x)[:3] == 'TFI' else 'ADCP后')
    
    # 新增用于计算变更次数列
    df['用于计算变更次数'] = df.apply(lambda row: 
        0.5 if (row['变更次数'] > 1 and 
              len(str(row['更改号自'])) > 5 and 
              len(str(row['更改号至'])) > 5 and
              (df[(df['更改号至'] == row['更改号自'])].empty or 
               df[(df['更改号自'] == row['更改号至'])].empty))
        else (1 if (row['变更次数'] > 1 and 
                  len(str(row['更改号自'])) > 5 and 
                  len(str(row['更改号至'])) > 5)
             else (1 if row['变更次数'] == 0 else 0.5)), 
        axis=1)
    
    # 统计各产品组ADCP前/后的变更次数
    if '产品组描述' in df.columns:
        group_stats = df.groupby(['产品组描述', '是否在ADCP前'])['用于计算变更次数'].sum().unstack()
        group_stats.columns = ['ADCP前变更次数', 'ADCP后变更次数']
        # 对空值补0
        group_stats = group_stats.fillna(0)
        # 计算总变更次数
        group_stats['总变更次数'] = group_stats['ADCP前变更次数'] + group_stats['ADCP后变更次数']
        
        # 将统计结果匹配回原表
        df = df.merge(group_stats, how='left', on='产品组描述')
    else:
        print("警告: 缺少'产品组描述'列，无法计算分组变更次数")

    # # 新增：统计每个产品组下标准型号的去重计数
    # if '产品组描述' in df.columns and '标准型号' in df.columns:
    #     # 计算每个产品组下不重复标准型号的数量
    #     group_model_counts = df.groupby('产品组描述')['标准型号'].nunique()
    #     # 映射到原始DataFrame
    #     df['产品组型号数'] = df['产品组描述'].map(group_model_counts)
    # else:
    #     print("警告: 缺少'产品组'或'标准型号'列，无法计算产品组型号数")

    # 新增：统计每个产品组下父级物料编码的去重计数
    if '产品组描述' in df.columns and '父项物料编码' in df.columns:
        # 计算每个产品组下不重复父级物料编码的数量
        group_parent_counts = df.groupby('产品组描述')['父项物料编码'].nunique()
        # 映射到原始DataFrame
        df['产品组父级物料数'] = df['产品组描述'].map(group_parent_counts)
    else:
        print("警告: 缺少'产品组'或'父项物料编码'列，无法计算产品组父级物料数")

    # 统计各父级物料编码ADCP前/后的变更次数
    if '父项物料编码' in df.columns:
        # 按父级物料编码和ADCP前后分组统计
        parent_stats = df.groupby(['父项物料编码', '是否在ADCP前'])['用于计算变更次数'].sum().unstack()
        parent_stats.columns = ['父项ADCP前变更次数', '父项ADCP后变更次数']
        # 对空值补0
        parent_stats = parent_stats.fillna(0)
        # 计算总变更次数
        parent_stats['父项总变更次数'] = parent_stats['父项ADCP前变更次数'] + parent_stats['父项ADCP后变更次数']
        
        # 新增变更情况种类判断
        parent_stats['变更情况种类'] = '无变更'
        parent_stats.loc[parent_stats['父项ADCP前变更次数'] > 0, '变更情况种类'] = '仅ADCP前变更'
        parent_stats.loc[parent_stats['父项ADCP后变更次数'] > 0, '变更情况种类'] = '仅ADCP后变更'
        parent_stats.loc[(parent_stats['父项ADCP前变更次数'] > 0) & 
                        (parent_stats['父项ADCP后变更次数'] > 0), '变更情况种类'] = 'ADCP前后均变更'
        
        # 将统计结果匹配回原表
        df = df.merge(parent_stats, how='left', on='父项物料编码')

        # 新增：统计每个产品组下不同变更情况种类的父级物料去重计数
        if '产品组描述' in df.columns:
            # 计算每个产品组下不同变更情况种类的父级物料数量
            group_change_counts = df.groupby(['产品组描述', '变更情况种类'])['父项物料编码'].nunique().unstack()
            group_change_counts = group_change_counts.fillna(0).astype(int)
            
            # 重命名列名
            group_change_counts.columns = [f'{col}父项物料数（产品组汇总）' for col in group_change_counts.columns]
            
            # 将统计结果匹配回原表
            df = df.merge(group_change_counts, how='left', on='产品组描述')
    else:
        print("警告: 缺少'父项物料编码'列，无法计算分组变更次数")
    
    # 新增：统计每个产品组下变更次数≥2的子项物料去重计数
    if '产品组描述' in df.columns and '变更次数' in df.columns:
        # 筛选变更次数≥2的子项物料
        high_change_materials = df[df['变更次数'] >= 2]
        # 计算每个产品组下符合条件的子项物料去重计数
        group_high_change_counts = high_change_materials.groupby('产品组描述')['最原始子项物料'].nunique()
        # 映射到原始DataFrame
        df['产品组高变更子项物料数'] = df['产品组描述'].map(group_high_change_counts).fillna(0).astype(int)
        
        # 新增：统计ADCP前高变更子项物料数
        adcp_before_high = high_change_materials[high_change_materials['是否在ADCP前'] == 'ADCP前']
        group_adcp_before_high = adcp_before_high.groupby('产品组描述')['最原始子项物料'].nunique()
        df['ADCP前高变更子项物料数'] = df['产品组描述'].map(group_adcp_before_high).fillna(0).astype(int)
        
        # 新增：统计ADCP后高变更子项物料数
        adcp_after_high = high_change_materials[high_change_materials['是否在ADCP前'] == 'ADCP后']
        group_adcp_after_high = adcp_after_high.groupby('产品组描述')['最原始子项物料'].nunique()
        df['ADCP后高变更子项物料数'] = df['产品组描述'].map(group_adcp_after_high).fillna(0).astype(int)
        
        # 新增：计算ADCP前变更均值
        df['ADCP前变更物料均值'] = df['ADCP前高变更子项物料数'] / (
            df['ADCP前后均变更父项物料数（产品组汇总）'] + 
            df['仅ADCP前变更父项物料数（产品组汇总）']
        ).replace([np.inf, -np.inf], 0).fillna(0)
        
        # 新增：计算ADCP后变更均值
        df['ADCP后变更物料均值'] = df['ADCP后高变更子项物料数'] / (
            df['ADCP前后均变更父项物料数（产品组汇总）'] + 
            df['仅ADCP后变更父项物料数（产品组汇总）']
        ).replace([np.inf, -np.inf], 0).fillna(0)
        
        # 新增：计算整体变更均值
        df['整体变更物料均值'] = df['产品组高变更子项物料数'] / df['产品组父级物料数'].replace([np.inf, -np.inf], 0).fillna(0)

        # 新增：计算ADCP前变更均值
        df['ADCP前变更次数均值'] = df['ADCP前变更次数'] / (
            df['ADCP前后均变更父项物料数（产品组汇总）'] + 
            df['仅ADCP前变更父项物料数（产品组汇总）']
        ).replace([np.inf, -np.inf], 0).fillna(0)
        
        # 新增：计算ADCP后变更均值
        df['ADCP后变更次数均值'] = df['ADCP后变更次数'] / (
            df['ADCP前后均变更父项物料数（产品组汇总）'] + 
            df['仅ADCP后变更父项物料数（产品组汇总）']
        ).replace([np.inf, -np.inf], 0).fillna(0)
        
        # 新增：计算整体变更均值
        df['整体变更次数均值'] = df['总变更次数'] / df['产品组父级物料数'].replace([np.inf, -np.inf], 0).fillna(0)


    else:
        print("警告: 缺少'产品组描述'或'变更次数'列，无法计算高变更子项物料数")

    # 保存带有变更次数的新Excel文件
    output_file = file_path.replace('.xlsx', '_带变更次数-使用最终子项物料.xlsx')
    df.to_excel(output_file, index=False)
    
    return df

if __name__ == "__main__":
    # 替换为你的Excel文件路径
    excel_file = r"D:\000物料报表\202504\物料变更\ECN与生命周期合并结果.xlsx"
    result = count_material_changes(excel_file)
    #分组计算的逻辑是按照整机来的，所以要把最终父项物料编码的列名改成父项物料编码（手动）（代码里面这样写的）
    # 输出变更次数统计结果
    print("处理完成，已生成带变更次数的Excel文件")
