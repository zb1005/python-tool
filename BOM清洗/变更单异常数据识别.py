import pandas as pd
from functools import partial

# 定义常量
RELATION_FILE_PATH = r'C:\Users\zhangbon\Desktop\制造研发变更对照表.XLSX'
OUTPUT_FILE_PATH = r'C:\Users\zhangbon\Desktop\升版信息及差异.xlsx'
VERSION_ORDER = {'0':0, 'A':1, 'B':2, 'C':3, 'D':4, 'E':5, 'F':6, 'G':7, 'H':8, 'I':9, 'J':10, 'K':11, 'L':12}

# 数据加载函数
def load_excel(file_path):
    """加载Excel文件并返回DataFrame"""
    return pd.read_excel(file_path)

# 数据筛选更改号自，至非空函数
def filter_empty_change_numbers(df):
    """筛选出更改号自和更改号至不同时为空的记录"""
    return df[~(df['更改号自'].isna() & df['更改号至'].isna())]

# 研发ECN匹配函数
def create_ecn_matcher(relation_df):
    """创建ECN匹配函数"""
    relation_dict = dict(zip(relation_df['制造ECN号'], relation_df['研发ECN']))
    
    def match_ecn(row):
        if pd.notna(row['更改号至']) and len(str(row['更改号至']))>2:
            return relation_dict.get(row['更改号至'], 0)
        elif pd.notna(row['更改号自']) and len(str(row['更改号自']))>2:
            return relation_dict.get(row['更改号自'], 0)
        return None
    
    return match_ecn

# 版本特征提取函数
def extract_version_features(df):
    """提取子级物料编码的版本特征"""
    df_copy = df.copy()
    df_copy['子级前12位'] = df_copy['子级物料编码'].astype(str).str[:12]
    df_copy['子级最后一位'] = df_copy['子级物料编码'].astype(str).str[-1]
    df_copy['版本顺序'] = df_copy['子级最后一位'].map(VERSION_ORDER)
    return df_copy

# 升版分析函数
def analyze_upgrades(grouped_data):
    """分析升版情况，返回连续升版、非连续升版和差异记录"""
    upgrade_records = []
    non_continuous_records = []
    diff_records = []
    
    for _, group in grouped_data:
        valid_group = group.dropna(subset=['版本顺序']).sort_values('版本顺序')
        if len(valid_group) >= 2:
            version_diff = valid_group['版本顺序'].diff().dropna()
            is_continuous = (version_diff == 1).all()
            
            if is_continuous:
                upgrade_records.append(valid_group)
            else:
                non_continuous_records.append(valid_group)
            
            # 检查更改号连续性
            for i in range(1, len(valid_group)):
                prev, current = valid_group.iloc[i-1], valid_group.iloc[i]
                if str(prev['更改号至']).strip() != str(current['更改号自']).strip():
                    diff_records.append({
                        '工厂': prev['工厂'],
                        '父级物料编码': prev['父级物料编码'],
                        '旧版子级物料': prev['子级物料编码'],
                        '新版子级物料': current['子级物料编码'],
                        '旧版更改号至': prev['更改号至'],
                        '新版更改号自': current['更改号自']
                    })
    
    return (
        pd.concat(upgrade_records) if upgrade_records else None,
        pd.concat(non_continuous_records) if non_continuous_records else None,
        pd.DataFrame(diff_records) if diff_records else None
    )

# 结果保存函数
def save_results(continuous_df, non_continuous_df, diff_df, output_path):
    """保存分析结果到Excel文件"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if continuous_df is not None:
            continuous_df.to_excel(writer, sheet_name='升版数据', index=False)
        if non_continuous_df is not None:
            non_continuous_df.to_excel(writer, sheet_name='非连续升版数据', index=False)
        if diff_df is not None:
            diff_df.to_excel(writer, sheet_name='升版差异数据', index=False)
    return output_path

# 主函数
def main(input_file_path=r"C:\Users\zhangbon\Desktop\PLM-MBOM数据核对表-20250618.xlsx"):
    """主函数：执行完整的数据处理流程"""
    # 加载数据
    raw_data = load_excel(input_file_path)
    relation_data = load_excel(RELATION_FILE_PATH)
    
    # 数据处理管道
    filtered_data = filter_empty_change_numbers(raw_data)
    match_ecn = create_ecn_matcher(relation_data)
    data_with_ecn = filtered_data.assign(研发ECN=filtered_data.apply(match_ecn, axis=1))
    data_with_features = extract_version_features(data_with_ecn)
    data_with_features.to_excel(r"C:\Users\zhangbon\Desktop\PLM-MBOM-有效-存在变更-匹配了研发ecn.xlsx", index=False)
    
    # 分组分析
    grouped = data_with_features.groupby(['工厂', '父级物料编码', '子级前12位'])
    continuous, non_continuous, diffs = analyze_upgrades(grouped)
    
    # 保存结果
    output_path = save_results(continuous, non_continuous, diffs, OUTPUT_FILE_PATH)
    return f'升版信息处理完成，结果已保存至"{output_path}"'

if __name__ == "__main__":
    main()
