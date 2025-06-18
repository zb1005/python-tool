import os
import pandas as pd

def get_filenames_to_excel(folder_path, output_file):
    """
    读取文件夹下所有文件名并保存到Excel
    :param folder_path: 文件夹路径
    :param output_file: 输出的Excel文件名
    """
    # 获取文件夹下所有文件名
    filenames = os.listdir(folder_path)
    
    # 创建DataFrame
    df = pd.DataFrame(filenames, columns=['文件名'])
    
    # 保存到Excel
    df.to_excel(output_file, index=False)
    print(f"文件名已保存到 {output_file}")

if __name__ == "__main__":
    # 示例用法
    folder_path = r"C:\Users\zhangbon\Desktop\数据产权\2024年方太数据知识产权登记证书"
    output_file = r"C:\Users\zhangbon\Desktop\数据产权\2024年方太数据知识产权登记证书\1.xlsx"
    get_filenames_to_excel(folder_path, output_file)