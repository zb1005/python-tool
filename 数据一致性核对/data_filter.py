"""
数据筛选规则模块
功能：根据目标文件名解析系统类型和数据类型，对源数据进行筛选
"""

import pandas as pd
import re
import os
from typing import Dict, List, Callable, Optional


def parse_filename_info(filename: str) -> Dict[str, str]:
    """
    解析文件名，提取系统类型和数据类型
    文件名格式示例：ERP-物料.xlsx
    
    Args:
        filename: 文件名（带或不带扩展名）
        
    Returns:
        Dict: 包含系统类型和数据类型的字典
    """
    # 去除文件扩展名
    basename = os.path.splitext(filename)[0]
    
    # 使用正则表达式匹配 "系统-数据类型" 格式
    pattern = r'^([^-]+)-(.+)$'
    match = re.match(pattern, basename)
    
    if match:
        system_type = match.group(1).strip()
        data_type = match.group(2).strip()
        return {
            'system': system_type,
            'data_type': data_type
        }
    else:
        # 如果不符合标准格式，尝试其他解析方式
        parts = basename.split('-')
        if len(parts) >= 2:
            return {
                'system': parts[0].strip(),
                'data_type': '-'.join(parts[1:]).strip()
            }
        else:
            # 如果无法解析，返回默认值
            return {
                'system': '未知系统',
                'data_type': '未知类型'
            }


def get_filter_rules(system: str, data_type: str, key_column: str) -> List[Callable]:
    """
    根据系统类型和数据类型获取筛选规则列表
    
    Args:
        system: 系统类型（如ERP、PLM等）
        data_type: 数据类型（如物料、客户等）
        key_column: 主键列名（已简化，不再依赖主键）
        
    Returns:
        List[Callable]: 筛选规则函数列表
    """
    
    # 物料数据类型的筛选规则（按系统分类）
    if data_type == "物料":
        if system == "ECS":
            # ECS物料筛选规则：10-17开头，且工厂为1000-1005，且生产仓储地点或外部采购仓库地点不为空或0
            return [
                lambda df: df[
                    df.get('物料编码', '').str.startswith(('10', '11', '12', '13', '14', '15', '16', '17'), na=False) & 
                    df.get('工厂', '').isin(['1000', '1001', '1002', '1003', '1004', '1005']) & 
                    ((df.get('生产仓储地点', '').notna() & (df.get('生产仓储地点', '') != '') & (df.get('生产仓储地点', '') != '0')) |
                     (df.get('外部采购仓库地点', '').notna() & (df.get('外部采购仓库地点', '') != '') & (df.get('外部采购仓库地点', '') != '0')))
                ]
            ]
        elif system in ["APS", "PCT", "SRM", "MES", "WMS原材料"]:
            # APS/PCT/SRM/MES/WMS原材料物料筛选规则：10-15开头
            return [
                lambda df: df[df.get('物料编码', '').str.startswith(('10', '11', '12', '13', '14', '15'), na=False)]
            ]
        elif system == "WMS配件":
            # WMS配件物料筛选规则：11,12,13,15开头且是否服务配件=X
            return [
                lambda df: df[
                    df.get('物料编码', '').str.startswith(('11', '12', '13', '15'), na=False) & 
                    (df.get('是否服务配件', '') == 'X')
                ]
            ]
        elif system == "TMS":
            # TMS物料筛选规则：10,11,1512,1513开头
            return [
                lambda df: df[df.get('物料编码', '').str.startswith(('10', '11', '1512', '1513'), na=False)]
            ]
        elif system == "WMS立库":
            # WMS立库物料筛选规则：10,1512,1513,159902开头，或是关键件、附件、独立包装
            return [
                lambda df: df[
                    df.get('物料编码', '').str.startswith(('10', '1512', '1513', '159902'), na=False) |
                    (df.get('是否关键件', '') == 'X') |
                    (df.get('是否附件', '') == 'X') |
                    (df.get('是否独立包装', '') == 'X')
                ]
            ]
        elif system == "CSM":
            # CSM物料筛选规则：10开头且项目阶段为特定阶段，或90,1703开头，或是服务配件、服务虚拟件
            project_stages = ['TR5', '验证阶段', 'TR6', 'ADCP', '发布阶段', 'TR7', '结项阶段']
            return [
                lambda df: df[
                    (df.get('物料编码', '').str.startswith('10', na=False) & df.get('项目阶段', '').isin(project_stages)) |
                    df.get('物料编码', '').str.startswith(('90', '1703'), na=False) |
                    (df.get('是否服务配件', '') == 'X') |
                    (df.get('是否服务虚拟件', '') == 'X')
                ]
            ]
        elif system == "CEM":
            # CEM物料筛选规则：10开头且项目阶段为特定阶段，或90,1703,1512,1513开头，或是服务配件、服务虚拟件，或11开头且独立包装
            project_stages = ['TR5', '验证阶段', 'TR6', 'ADCP', '发布阶段', 'TR7', '结项阶段']
            return [
                lambda df: df[
                    (df.get('物料编码', '').str.startswith('10', na=False) & df.get('项目阶段', '').isin(project_stages)) |
                    df.get('物料编码', '').str.startswith(('90', '1703', '1512', '1513'), na=False) |
                    (df.get('是否服务配件', '') == 'X') |
                    (df.get('是否服务虚拟件', '') == 'X') |
                    (df.get('物料编码', '').str.startswith('11', na=False) & (df.get('是否独立包装', '') == 'X'))
                ]
            ]
        elif system == "EPS":
            # EPS物料筛选规则：15,16开头
            return [
                lambda df: df[df.get('物料编码', '').str.startswith(('15', '16'), na=False)]
            ]
        elif system == "电商":
            # 电商物料筛选规则：10开头且项目阶段为特定阶段，或1512,1703开头
            project_stages = ['TR5', '验证阶段', 'TR6', 'ADCP', '发布阶段', 'TR7', '结项阶段']
            return [
                lambda df: df[
                    (df.get('物料编码', '').str.startswith('10', na=False) & df.get('项目阶段', '').isin(project_stages)) |
                    df.get('物料编码', '').str.startswith(('1512', '1703'), na=False)
                ]
            ]
        elif system == "LTC":
            # LTC物料筛选规则：10,17020000开头，或是服务配件、独立包装
            return [
                lambda df: df[
                    df.get('物料编码', '').str.startswith(('10', '17020000'), na=False) |
                    (df.get('是否服务配件', '') == 'X') |
                    (df.get('是否独立包装', '') == 'X')
                ]
            ]
        elif system == "数据中台":
            # 数据中台物料筛选规则：10,11,1512,1513开头，或12,13,15开头且服务配件
            return [
                lambda df: df[
                    df.get('物料编码', '').str.startswith(('10', '11', '1512', '1513'), na=False) |
                    (df.get('物料编码', '').str.startswith(('12', '13', '15'), na=False) & (df.get('是否服务配件', '') == 'X'))
                ]
            ]
        elif system == "FIKS平台":
            # FIKS平台物料筛选规则：10开头
            return [
                lambda df: df[df.get('物料编码', '').str.startswith('10', na=False)]
            ]
        elif system == "DCS":
            # DCS物料筛选规则：10,11,1512,1513开头
            return [
                lambda df: df[df.get('物料编码', '').str.startswith(('10', '11', '1512', '1513'), na=False)]
            ]
        elif system == "DMS":
            # DMS物料筛选规则：10开头且项目阶段为特定阶段，或91,1512,1513,1701,1702,159902,159903开头，或独立包装
            project_stages = ['TR5', '验证阶段', 'TR6', 'ADCP', '发布阶段', 'TR7', '结项阶段']
            return [
                lambda df: df[
                    (df.get('物料编码', '').str.startswith('10', na=False) & df.get('项目阶段', '').isin(project_stages)) |
                    df.get('物料编码', '').str.startswith(('91', '1512', '1513', '1701', '1702', '159902', '159903'), na=False) |
                    (df.get('是否独立包装', '') == 'X')
                ]
            ]
    
    # 客户数据类型的筛选规则
    elif data_type == "客户":
        if system == "ERP":
            # ERP客户筛选规则：客户编码长度为10
            return [
                lambda df: df[df.get('客户编码', '').str.len() == 10]
            ]
    
    # 如果没有匹配的规则，返回空列表
    return []


def apply_filter_rules(source_df: pd.DataFrame, target_filename: str, key_columns: List[str]) -> pd.DataFrame:
    """
    根据目标文件名信息对源数据进行筛选
    
    Args:
        source_df: 源数据DataFrame
        target_filename: 目标文件名
        key_columns: 主键列名列表（已简化，不再依赖主键）
        
    Returns:
        pd.DataFrame: 筛选后的数据
    """
    # 解析目标文件名信息
    file_info = parse_filename_info(target_filename)
    system = file_info['system']
    data_type = file_info['data_type']
    
    print(f"🔍 应用筛选规则 - 系统: {system}, 数据类型: {data_type}")
    
    # 获取筛选规则（不再依赖主键列）
    primary_key = key_columns[0] if key_columns else "默认主键"
    filter_rules = get_filter_rules(system, data_type, primary_key)
    
    if not filter_rules:
        print(f"ℹ️ 未找到针对系统 '{system}' 和数据类型 '{data_type}' 的筛选规则")
        return source_df
    
    # 应用筛选规则（现在每个系统只有一个合并的规则函数）
    original_count = len(source_df)
    
    for i, rule in enumerate(filter_rules, 1):
        try:
            result_df = rule(source_df.copy())
            filtered_count = len(result_df)
            
            if not result_df.empty:
                print(f"✅ 筛选完成: {original_count} → {filtered_count} 条记录 (过滤 {original_count - filtered_count} 条)")
                return result_df
            else:
                print(f"⚠️ 规则{i}未匹配到任何记录")
                return source_df.iloc[0:0]  # 返回空DataFrame
        except Exception as e:
            print(f"⚠️ 规则{i}执行失败: {str(e)}")
            return source_df.iloc[0:0]  # 返回空DataFrame
    
    # 如果没有规则，返回原始数据
    return source_df


def filter_data_for_comparison(source_df: pd.DataFrame, target_paths: List[str], 
                              key_columns: List[str]) -> Dict[str, pd.DataFrame]:
    """
    为批量比对准备筛选后的源数据
    
    Args:
        source_df: 原始源数据
        target_paths: 目标文件路径列表
        key_columns: 主键列名列表
        
    Returns:
        Dict: 键为目标文件名，值为对应的筛选后源数据
    """
    filtered_data = {}
    
    for target_path in target_paths:
        target_filename = os.path.basename(target_path)
        
        # 对每个目标文件应用独立的筛选
        filtered_df = apply_filter_rules(source_df.copy(), target_filename, key_columns)
        filtered_data[target_filename] = filtered_df
        
        print(f"📊 {target_filename}: 筛选后记录数 = {len(filtered_df)}")
    
    return filtered_data


# # 示例：添加新规则的辅助函数
# def add_custom_rule(system: str, data_type: str, key_column: str, 
#                    rule_func: Callable, rule_name: str = "自定义规则"):
#     """
#     添加自定义筛选规则（示例函数，实际使用时需要修改get_filter_rules函数）
    
#     Args:
#         system: 系统类型
#         data_type: 数据类型
#         key_column: 主键列名
#         rule_func: 筛选规则函数
#         rule_name: 规则名称
#     """
#     # 在实际使用中，您需要修改get_filter_rules函数来包含这些自定义规则
#     # 这里只是一个示例接口
#     print(f"📝 添加规则: {system}-{data_type}-{key_column} - {rule_name}")


# if __name__ == "__main__":
#     # 测试文件名解析
#     test_files = [
#         "ERP-物料.xlsx",
#         "PLM-物料.xlsx", 
#         "CRM-客户数据.xlsx",
#         "invalid_filename.txt"
#     ]
    
#     for filename in test_files:
#         info = parse_filename_info(filename)
#         print(f"{filename} → 系统: {info['system']}, 数据类型: {info['data_type']}")