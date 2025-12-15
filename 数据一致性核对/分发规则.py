import pandas as pd
def distribute_rule(source_df,target_df):
    """
    分发规则函数 - 根据源文件的物料组描述，将目标文件的行分发到对应的物料组描述中
    """
    # 引入物料组（大系列）
    df_bigtype = pd.read_excel(r'C:\Users\zhangbon\Desktop\物料组对照（大系列）.xlsx')
    df_bigtype['物料组'] = df_bigtype['物料组'].astype(str)
    bigtype_map = dict(zip(df_bigtype['物料组'], df_bigtype['物料组描述']))
    # bigtype_map

















