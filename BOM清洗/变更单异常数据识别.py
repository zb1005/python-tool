import pandas as pd
from functools import partial

# 定义常量
RELATION_FILE_PATH = r'C:\Users\zhangbon\Desktop\000000我的事项\2025-06\研发变更错误数据识别\制造研发变更对照表.XLSX'
OUTPUT_FILE_PATH = r'C:\Users\zhangbon\Desktop\000000我的事项\2025-06\研发变更错误数据识别\升版信息及差异.xlsx'
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
    diff_研发_records = []

    
    for _, group in grouped_data:
        # 按工厂、父级物料编码、子级物料编码、计数器升序排序
        valid_group = group.dropna(subset=['版本顺序', '子级前12位', '子级最后一位'])
        valid_group = valid_group.sort_values(by=['工厂', '父级物料编码', '子级物料编码', '计数器'])
        
        if len(valid_group) >= 2:
            for i in range(1, len(valid_group)):
                prev, current = valid_group.iloc[i-1], valid_group.iloc[i]
                
                # 检查子级物料前12位是否相同
                if prev['子级前12位'] == current['子级前12位']:
                    # 最后一位相同的情况
                    if prev['子级最后一位'] == current['子级最后一位']:
                        # 仅检查更改号连续性
                        if str(prev['更改号至']).strip() != str(current['更改号自']).strip():
                            diff_records.append({
                                '工厂': prev['工厂'],
                                '父级物料编码': prev['父级物料编码'],
                                '旧版子级物料': prev['子级物料编码'],
                                '新版子级物料': current['子级物料编码'],
                                '旧版更改号至': prev['更改号至'],
                                '新版更改号自': current['更改号自'],
                                '差异类型': '同版本更改号不匹配'
                            })
                    # 最后一位不同的情况
                    else:
                        # 检查是否连续升版
                        version_diff = current['版本顺序'] - prev['版本顺序']
                        is_continuous = (version_diff == 1)
                        
                        # 检查更改号连续性
                        change_id_match = (str(prev['更改号至']).strip() == str(current['更改号自']).strip())
                        
                        # 记录升版类型
                        if is_continuous:
                            upgrade_records.append(pd.concat([prev.to_frame().T, current.to_frame().T]))
                            if not change_id_match:
                                diff_records.append({
                                    '工厂': prev['工厂'],
                                    '父级物料编码': prev['父级物料编码'],
                                    '旧版子级物料': prev['子级物料编码'],
                                    '新版子级物料': current['子级物料编码'],
                                    '旧版更改号至': prev['更改号至'],
                                    '新版更改号自': current['更改号自'],
                                    '差异类型': '连续升版更改号不匹配'
                                })
                        else:
                            non_continuous_records.append(pd.concat([prev.to_frame().T, current.to_frame().T]))
                            if not change_id_match:
                                diff_records.append({
                                    '工厂': prev['工厂'],
                                    '父级物料编码': prev['父级物料编码'],
                                    '旧版子级物料': prev['子级物料编码'],
                                    '新版子级物料': current['子级物料编码'],
                                    '旧版更改号至': prev['更改号至'],
                                    '新版更改号自': current['更改号自'],
                                    '差异类型': '非连续升版更改号不匹配'
                                })
    #找出upgrade_records中同一子级物料编码，研发ECN不同的项
    upgrade_records_df=pd.concat(upgrade_records) if upgrade_records else None
    for _,group in upgrade_records_df.groupby(['子级物料编码','研发ECN']):
        if len(group['研发ECN'].unique())>1:
            diff_研发_records.append(group)
    
    return (
        pd.concat(upgrade_records) if upgrade_records else None,
        pd.concat(non_continuous_records) if non_continuous_records else None,
        pd.DataFrame(diff_records) if diff_records else None,
        pd.DataFrame(diff_研发_records) if diff_研发_records else None
    )   

# 结果保存函数
def identify_ecn_inconsistencies(data):
    """识别同一父级物料编码和子级物料编码下研发ECN值不同的项"""
    # 确保必要的列存在
    required_cols = ['父级物料编码', '子级物料编码', '研发ECN']
    if not set(required_cols).issubset(data.columns):
        missing = [col for col in required_cols if col not in data.columns]
        raise ValueError(f"数据缺少必要的列: {missing}")

    # 筛选出同一父级物料编码和子级物料编码下研发ECN值不同的项
    ecn_inconsistencies = data.groupby(['父级物料编码', '子级物料编码']).filter(lambda x: len(x['研发ECN'].unique()) > 1)
    return ecn_inconsistencies
    

def save_results(continuous_df, non_continuous_df, diff_df,diff_研发_df, output_path):

    """保存分析结果到Excel文件"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if continuous_df is not None:
            continuous_df.to_excel(writer, sheet_name='升版数据', index=False)
        if non_continuous_df is not None:
            non_continuous_df.to_excel(writer, sheet_name='非连续升版数据', index=False)
        if diff_df is not None:
            diff_df.to_excel(writer, sheet_name='差异数据', index=False)
        if diff_研发_df is not None:
            diff_研发_df.to_excel(writer, sheet_name='差异数据-研发ECN', index=False)
    return output_path

# 主函数
def main(input_file_path=r"C:\Users\zhangbon\Desktop\000000我的事项\2025-06\研发变更错误数据识别\PLM-MBOM数据核对表-20250618.xlsx"):
    """主函数：执行完整的数据处理流程"""
    # 加载数据
    raw_data = load_excel(input_file_path)
    relation_data = load_excel(RELATION_FILE_PATH)
    print("数据加载完成")
    # 数据处理管道
    filtered_data = filter_empty_change_numbers(raw_data)
    match_ecn = create_ecn_matcher(relation_data)
    data_with_ecn = filtered_data.assign(研发ECN=filtered_data.apply(match_ecn, axis=1))
    data_with_features = extract_version_features(data_with_ecn)
    data_with_features.to_excel(r"C:\Users\zhangbon\Desktop\000000我的事项\2025-06\研发变更错误数据识别\PLM-MBOM-有效-存在变更-匹配了研发ecn.xlsx", index=False)
    # 识别ECN不一致项
    data_with_ecn_check = identify_ecn_inconsistencies(data_with_features)
    data_with_ecn_check.to_excel(r"C:\Users\zhangbon\Desktop\000000我的事项\2025-06\研发变更错误数据识别\同一父级物料编码下的同一子级物料编码存在不同研发ECN.xlsx", index=False)
    
    # 分组分析
    grouped = data_with_features.groupby(['工厂', '父级物料编码', '子级前12位'])
    # 传入df参数并接收返回的四个值
    continuous, non_continuous, diffs, diffs_研发 = analyze_upgrades(grouped)
    
    # 保存结果
    output_path = save_results(continuous, non_continuous, diffs,diffs_研发, OUTPUT_FILE_PATH)
    return f'升版信息处理完成，结果已保存至"{output_path}"'

if __name__ == "__main__":
    main()
