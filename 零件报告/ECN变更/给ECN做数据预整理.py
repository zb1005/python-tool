import pandas as pd
import os

###准备工作有这几个文档1、整机特殊采购类.XLSX 2、物料生命周期-产品线-组.XLSX 3、仅ecn记录.XLSX 4、筛后整机BOM清单.xlsx


# 处理第一个Excel文件
def process_first_excel(file_path):
    df = pd.read_excel(file_path)
    # 筛选工厂为1000且特殊采购类不为空的行
    filtered_df = df[(df['工厂'] == 1000) & (df['特殊采购类'].notna())]
    return filtered_df

# 处理第二个Excel文件
def process_second_excel(file_path):
    df = pd.read_excel(file_path)
    # 定义需要保留的生命周期值
    valid_status = ['开发', '样机', '小批量', '量产', '退市预警', '停止销售']
    # 筛选符合条件的行
    filtered_df = df[(df['产品生命周期状态'].isin(valid_status))&(df["国内/海外"]==20)]
    return filtered_df

def merge_bom_files(bom_files):
    """
    合并所有BOM文件，先筛选再合并
    Args:
        bom_files: BOM文件路径列表
    Returns:
        合并后的DataFrame
    """
    # 读取ECN记录
    ecn_df = pd.read_excel(r"D:\000物料报表\202504\物料变更\仅ecn记录.XLSX")
    ecn_list = ecn_df['制造ECN号'].tolist()
    
    # 新增制造ECN号和研发ECN的对照关系字典
    ecn_mapping = dict(zip(ecn_df['制造ECN号'], ecn_df['研发ECN']))
    
    frames = []
    for file in bom_files:
        try:
            df = pd.read_excel(file)
            # 先筛选当前BOM文件中符合ECN条件的记录
            filtered = df[
                (df['更改号自'].isin(ecn_list)) |
                (df['更改号至'].isin(ecn_list))
            ]
            
            # 添加研发ECN匹配逻辑
            def match_rd_ecn(row):
                if len(str(row['更改号自'])) > 5 and len(str(row['更改号至'])) > 5:
                    return ecn_mapping.get(row['更改号至'], None)
                elif len(str(row['更改号自'])) > 5:
                    return ecn_mapping.get(row['更改号自'], None)
                elif len(str(row['更改号至'])) > 5:
                    return ecn_mapping.get(row['更改号至'], None)
                return None
            
            filtered['研发ECN'] = filtered.apply(match_rd_ecn, axis=1)
            # 保留前20列和新添加的研发ECN列
            frames.append(filtered.iloc[:, list(range(20)) + [len(filtered.columns)-1]])  # 保留前20列
            
        except Exception as e:
            print(f"警告：处理文件 {file} 时出错 - {str(e)}")
    
    if not frames:
        raise ValueError("没有有效的BOM文件可合并")
        
    merged = pd.concat(frames)
    merged.to_excel(r"C:\Users\zhangbon\Desktop\常用物料信息表\筛后全整机BOM.xlsx", index=False)
    return merged

def filter_by_ecn(merged_bom):
    """
    处理已合并的BOM数据（现在只需要与生命周期数据合并）
    Args:
        merged_bom: 已合并的BOM DataFrame
    Returns:
        最终合并结果
    """
    # 读取物料生命周期-特殊采购-筛后.xlsx
    lifecycle_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\常用物料信息表\物料生命周期-特殊采购-筛后.xlsx")
    
    # 合并两个表
    final_result = pd.merge(
        lifecycle_df,
        merged_bom,
        left_on='物料编码',
        right_on='最终父项物料编码',
        how='inner'
    )
    
    # 保存最终结果
    final_result.to_excel(r"C:\Users\zhangbon\Desktop\常用物料信息表\ECN与生命周期合并结果.xlsx", index=False)
    return final_result

# 主函数
def main():
    # 替换为实际文件路径
    first_excel_path = r"C:\Users\zhangbon\Desktop\常用物料信息表\整机特殊采购类.XLSX"
    second_excel_path = r"C:\Users\zhangbon\Desktop\常用物料信息表\物料生命周期-产品线-组.XLSX"
    
    # 处理文件
    first_result = process_first_excel(first_excel_path)
    second_result = process_second_excel(second_excel_path)
    
    # 依据物料编码拼接两个结果
    merged_result = pd.merge(
        first_result,
        second_result,
        on='物料编码',  # 假设两个文件都有物料编码列
        how='inner'    # 保留同时满足的记录
    )
    
    # 特殊采购类与工厂对照关系
    special_purchase_mapping = {
        'Z1': '1001', 'Z2': '1002', 'Z3': '1003', 'Z4': '1001',
        'Z5': '1002', 'Z6': '1003', 'Z7': '1005', 'Z8': '1005',
        'Z9': '1004', 'ZA': '1004'
    }
    
    # 特殊采购类与产品组对照关系
    product_group_mapping = {
        'Z1': '油烟机', 'ZZ': '不考核产品', 'Z2': '灶具', 'Z3': '消毒柜',
        'Z7': '热水器', 'ZT': '热水器', 'Z6': '烤箱', 'Z4': '微波炉',
        'Z5': '蒸箱', 'Z8': '水槽洗碗机', 'ZK': '蒸微', 'ZL': '蒸烤烹饪机',
        'ZM': '灶集成', 'ZP': '灶集成', 'ZN': '灶集成', 'ZQ': '蒸烤微烹饪机',
        'Y1': '不考核产品', 'Y4': '灶烤烹饪机', 'Z9': '不考核产品', 'ZA': '不考核产品',
        'ZD': '家用净水机', 'ZR': '不考核产品', 'ZU': '不考核产品', 'ZE': '不考核产品',
        'ZH': '不考核产品', 'ZJ': '不考核产品', 'ZS': '嵌入式洗碗机', 'ZV': '不考核产品',
        'ZX': '不考核产品', 'Y2': '不考核产品', 'Y3': '不考核产品', 'F1': '不考核产品'
    }
    
    # 新增两列
    merged_result['特殊采购对应工厂'] = merged_result['特殊采购类'].map(special_purchase_mapping)
    merged_result['产品组描述'] = merged_result['产品组'].map(product_group_mapping)
    
    # 保存结果
    merged_result.to_excel(r"C:\Users\zhangbon\Desktop\常用物料信息表\物料生命周期-特殊采购-筛后.xlsx", index=False)

    # 处理BOM文件
    bom_files = [
        r"C:\Users\zhangbon\Desktop\常用物料信息表\筛后1001整机BOM.xlsx",
        r"C:\Users\zhangbon\Desktop\常用物料信息表\筛后1002整机BOM.xlsx",
        r"C:\Users\zhangbon\Desktop\常用物料信息表\筛后1003整机BOM.xlsx",
        r"C:\Users\zhangbon\Desktop\常用物料信息表\筛后1004整机BOM.xlsx",
        r"C:\Users\zhangbon\Desktop\常用物料信息表\筛后1005整机BOM.xlsx"
    ]
    merged_bom = merge_bom_files(bom_files)

    # 将BOM-ecn匹配进去产品组信息（函数名没改）
    filter_by_ecn(merged_bom)
    print("处理完成！")

    
if __name__ == '__main__':
    main()