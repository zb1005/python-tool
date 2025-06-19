import pandas as pd
from tqdm import tqdm

def generate_composite_key(row, key_columns):
    """生成组合主键"""
    return '_'.join([
        str(row[col]).strip() if pd.notna(row[col]) and str(row[col]).strip() != ''
        else '<空值>'
        for col in key_columns
    ])

def read_excel_safely(file_path, key_columns):
    """安全读取Excel文件"""
    try:
        df = pd.read_excel(file_path, dtype=str, keep_default_na=False)
        df.columns = df.columns.str.strip()
        missing_keys = [col for col in key_columns if col not in df.columns]
        if missing_keys:
            raise ValueError(f"缺失主键列: {', '.join(missing_keys)}")
        df['_composite_key'] = df.apply(
            lambda x: generate_composite_key(x, key_columns),
            axis=1
        )
        return df
    except Exception as e:
        raise ValueError(f"文件读取失败: {str(e)}")

def compare_datasets(source_df, target_df, key_columns):
    """核心比对函数（字典优化版）"""
    # 生成主键并检查重复
    source_df['_composite_key'] = source_df.apply(
        lambda x: generate_composite_key(x, key_columns), axis=1)
    target_df['_composite_key'] = target_df.apply(
        lambda x: generate_composite_key(x, key_columns), axis=1)
    
    # 检查主键唯一性
    if source_df['_composite_key'].duplicated().any():
        dup_keys = source_df[source_df['_composite_key'].duplicated()]['_composite_key'].unique()
        raise ValueError(f"源数据中存在重复主键: {', '.join(dup_keys[:3])}...")
    if target_df['_composite_key'].duplicated().any():
        dup_keys = target_df[target_df['_composite_key'].duplicated()]['_composite_key'].unique()
        raise ValueError(f"目标数据中存在重复主键: {', '.join(dup_keys[:3])}...")
    
    # 生成主键并转换为字典
    source_df['_composite_key'] = source_df.apply(
        lambda x: generate_composite_key(x, key_columns), axis=1)
    target_df['_composite_key'] = target_df.apply(
        lambda x: generate_composite_key(x, key_columns), axis=1)
    
    # 转换为字典提高查找效率
    source_dict = source_df.set_index('_composite_key').to_dict('index')
    target_dict = target_df.set_index('_composite_key').to_dict('index')
    
    # 获取主键集合和共同列
    source_keys = set(source_dict.keys())
    target_keys = set(target_dict.keys())
    common_columns = list(set(source_df.columns) & set(target_df.columns) - {'_composite_key'})
    
    results = []

    # 处理仅存在于源数据的主键
    for key in tqdm(source_keys - target_keys, desc="处理仅源数据记录"):
        results.append({
            '主键状态': '仅存在于源文件',
            '组合主键': key,
            **{f"源_{col}": source_dict[key][col] for col in common_columns},
            **{f"目标_{col}": "" for col in common_columns}
        })

    # 处理仅存在于目标数据的主键
    for key in tqdm(target_keys - source_keys, desc="处理仅目标数据记录"):
        results.append({
            '主键状态': '仅存在于目标文件',
            '组合主键': key,
            **{f"源_{col}": "" for col in common_columns},
            **{f"目标_{col}": target_dict[key][col] for col in common_columns}
        })

    # 处理共同主键的数据差异
    for key in tqdm(source_keys & target_keys, desc="比对共同记录"):
        src_row = source_dict[key]
        tgt_row = target_dict[key]

        diff_details = {}
        row_data = {'主键状态': '数据一致', '组合主键': key}

        for col in common_columns:
            src_val = src_row[col]
            tgt_val = tgt_row[col]
            
            if src_val != tgt_val:
                diff_details[col] = {'源值': src_val, '目标值': tgt_val}
            
            row_data[f"源_{col}"] = src_val
            row_data[f"目标_{col}"] = tgt_val

        if diff_details:
            row_data['主键状态'] = f"发现{len(diff_details)}处差异"
            row_data['差异详情'] = str(diff_details)
            row_data['差异列名'] = ', '.join(diff_details.keys())

        results.append(row_data)

    return pd.DataFrame(results)

def simple_compare(source_path, target_path, key_columns, output_path):
    """简化版比对主程序"""
    try:
        source_df = read_excel_safely(source_path, key_columns)
        target_df = read_excel_safely(target_path, key_columns)
        
        result_df = compare_datasets(source_df, target_df, key_columns)
        result_df.to_excel(output_path, index=False)
        print(f"比对完成，结果已保存至: {output_path}")
    except Exception as e:
        print(f"处理失败: {str(e)}")

# 使用示例
if __name__ == "__main__":

    SOURCE_FILE = r"C:\Users\zhangbon\Desktop\MDM-CEM物料20250605.XLSX"
    TARGET_FILE = r"C:\Users\zhangbon\Desktop\CEM物料_20250609.xlsx"
    KEY_COLUMNS = ["物料编码"]  # 主键列
    # KEY_COLUMNS = ["工厂", "物料编码","组件"]
    OUTPUT_FILE = r"C:\Users\zhangbon\Desktop\EBOM是源--核对结果.xlsx"
    simple_compare(SOURCE_FILE, TARGET_FILE, KEY_COLUMNS, OUTPUT_FILE)