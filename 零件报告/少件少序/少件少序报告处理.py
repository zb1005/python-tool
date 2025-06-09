import pandas as pd

# 定义对照关系字典
MAPPING_DICT = {
    # 特殊采购类与工厂对照
    'factory_mapping': {
        'Z1': '1001', 'Z2': '1002', 'Z3': '1003', 'Z4': '1001',
        'Z5': '1002', 'Z6': '1003', 'Z7': '1005', 'Z8': '1005',
        'Z9': '1004', 'ZA': '1004'
    },
    
    # 文件路径配置
    'file_paths': {
        #要匹配进去整机的创建时间、特殊采购类、产品组代码、产品组、产品线，先筛选一下，然后再导出新和旧的bom数据再用此程序
        'first': r"D:\000物料报表\202504\新品少件少序\对标型号清单定期维护20250508.XLSX",
        'new_model': r"D:\000物料报表\202504\新品少件少序\新.XLSX",
        'old_model': r"D:\000物料报表\202504\新品少件少序\旧.XLSX"
    },
    
    # 需要保留的字段
    'required_columns': [
        '工厂', '最终父项物料编码', '最终父项名称', '父项物料编码',
        '父项物料名称', '加工工厂', '行项目类别', '项目号',
        '子项物料编码', '子项物料描述', '数量', '单位.1'
    ]
}

def process_excel_files():
    # 读取第一个Excel文件
    df1 = pd.read_excel(MAPPING_DICT['file_paths']['first'], sheet_name='S3-S4')
    
    # 建立物料编码和特殊采购类的对照关系
    material_purchase = df1[['物料编码', '特殊采购类']].drop_duplicates()
    
    # 建立对标型号-物料编码和特殊采购类-对标型号的对应关系
    model_mapping = df1[['对标型号-物料编码', '特殊采购类-对标型号']].drop_duplicates()
    
    # 结合特殊采购类和工厂的对照关系
    df1['工厂'] = df1['特殊采购类'].map(MAPPING_DICT['factory_mapping'])
    # 结合特殊采购类和工厂的对照关系
    df1['工厂-对照型号'] = df1['特殊采购类-对标型号'].map(MAPPING_DICT['factory_mapping'])
    # 结合特殊采购类和工厂的对照关系
    material_purchase['工厂'] = material_purchase['特殊采购类'].map(MAPPING_DICT['factory_mapping'])
    # 结合特殊采购类和工厂的对照关系
    model_mapping['工厂'] = model_mapping['特殊采购类-对标型号'].map(MAPPING_DICT['factory_mapping'])
    
    # 处理新型号文件
    new_model_df = pd.read_excel(MAPPING_DICT['file_paths']['new_model'])
    new_model_df['工厂'] = new_model_df['工厂'].astype(str)
    
    new_model_df = new_model_df.merge(
        material_purchase[['物料编码', '工厂']].rename(columns={'物料编码': '最终父项物料编码'}),
        left_on=['最终父项物料编码', '工厂'],
        right_on=['最终父项物料编码', '工厂'],
        how='inner'
    )
    # 处理新型号文件
    new_model_df = new_model_df[MAPPING_DICT['required_columns']]
    
    # 新型号数据聚合计算
    # 新型号数据聚合计算
    new_stats = new_model_df.groupby('最终父项物料编码').apply(
        lambda g: pd.Series({
            '子项11开头数量': g['子项物料编码'].astype(str).str.startswith('11').sum(),
            '子项13开头螺钉数量': g.loc[
                (g['子项物料编码'].astype(str).str.startswith('13')) & 
                (g['子项物料描述'].str.contains('螺钉')), 
                '数量'
            ].sum()
        })
    ).reset_index()
    
    # 将新型号统计结果匹配到df1
    df1 = df1.merge(
        new_stats,
        left_on='物料编码',
        right_on='最终父项物料编码',
        how='left'
    )
    
    # 处理老型号文件
    old_model_df = pd.read_excel(MAPPING_DICT['file_paths']['old_model'])
    old_model_df['工厂'] = old_model_df['工厂'].astype(str)
    
    old_model_df = old_model_df.merge(
        model_mapping[['对标型号-物料编码', '工厂']].rename(columns={'对标型号-物料编码': '最终父项物料编码'}),
        left_on=['最终父项物料编码', '工厂'],
        right_on=['最终父项物料编码', '工厂'],
        how='inner'
    )
    old_model_df = old_model_df[MAPPING_DICT['required_columns']]
    # 老型号数据聚合计算
    old_stats = old_model_df.groupby('最终父项物料编码').apply(
        lambda g: pd.Series({
            '子项11开头数量_对照型号': g['子项物料编码'].astype(str).str.startswith('11').sum(),
            '子项13开头螺钉数量_对照型号': g.loc[
                (g['子项物料编码'].astype(str).str.startswith('13')) & 
                (g['子项物料描述'].str.contains('螺钉')), 
                '数量'
            ].sum()
        })
    ).reset_index()
    
    # 将老型号统计结果匹配到material_purchase
    df1 = df1.merge(
        old_stats,
        left_on='对标型号-物料编码',
        right_on='最终父项物料编码',
        how='left'
    )
    
    #判断新型号与旧型号相比是否少件
    df1["是否少件"] = df1.apply(
        lambda row: "1" if row["子项11开头数量"] < row["子项11开头数量_对照型号"] else "0",
        axis=1
    )
    #判断新型号与旧型号相比是否少序
    df1["是否少序"] = df1.apply(
        lambda row: "1" if row["子项13开头螺钉数量"] < row["子项13开头螺钉数量_对照型号"] else "0",
        axis=1
    )
    #分组计算每个产品组的少件率和少序率再匹配进去
    df1['产品组少件率'] = df1.groupby('产品组')['是否少件'].transform(lambda x: (x == '1').sum() / x.count())
    #分组计算每个产品组的少序率
    df1['产品组少序率'] = df1.groupby('产品组')['是否少序'].transform(lambda x: (x == '1').sum() / x.count())
    
    # 保存结果
    df1.to_excel(r'D:\000物料报表\202504\新品少件少序\统计结果.xlsx', index=False)
    new_model_df.to_excel(r'D:\000物料报表\202504\新品少件少序\筛选后的新型号文件.xlsx', index=False)
    old_model_df.to_excel(r'D:\000物料报表\202504\新品少件少序\筛选后的老型号文件.xlsx', index=False)
    new_stats.to_excel(r'D:\000物料报表\202504\新品少件少序\新型号统计结果.xlsx', index=False)
    old_stats.to_excel(r'D:\000物料报表\202504\新品少件少序\老型号统计结果.xlsx', index=False)
    print("处理完成！")
    print(new_model_df.columns)
if __name__ == '__main__':
    process_excel_files()
