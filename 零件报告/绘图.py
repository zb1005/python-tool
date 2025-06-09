import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# 读取Excel文件
def read_excel_and_generate_wordcloud(file_path, word_column, count_column=None):
    # 读取Excel文件
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # 检查列是否存在
    if word_column not in df.columns:
        raise ValueError(f"列名 '{word_column}' 不存在于Excel文件中")
    if count_column and count_column not in df.columns:
        raise ValueError(f"列名 '{count_column}' 不存在于Excel文件中")
    
    # 获取词和计数数据
    if count_column:
        word_freq = {word: count for word, count in zip(df[word_column], df[count_column])}
        wordcloud = WordCloud(
            width=800,
            height=400,
            background_color='white',
            font_path='simhei.ttf'
        ).generate_from_frequencies(word_freq)
    else:
        text_data = ' '.join(df[word_column].astype(str).tolist())
        wordcloud = WordCloud(
            width=800, 
            height=400,
            background_color='white',
            font_path='simhei.ttf'
        ).generate(text_data)
    
    # 显示词云
    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.show()

# 使用示例
if __name__ == '__main__':
    excel_path = r'C:\Users\zhangbon\Desktop\1.xlsx'  # 替换为你的Excel文件路径
    word_column = '数据类型'  # 替换为词列名
    count_column = '计数'  # 替换为计数列名
    read_excel_and_generate_wordcloud(excel_path, word_column, count_column)