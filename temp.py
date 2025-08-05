import os
import re
# # 读取一个txt文件，并将每一行的内容使用；连接
# def read_and_join_lines(file_path):
#     try:
#         with open(file_path, 'r', encoding='utf-8') as file:
#             # 读取所有行并去除每行首尾的空白字符
#             lines = [line.strip() for line in file.readlines()]
#             # 使用分号连接所有行
#             return ';'.join(lines)
#     except FileNotFoundError:
#         return f"错误：文件 '{file_path}' 不存在"
#     except Exception as e:
#         return f"读取文件时发生错误：{str(e)}"

# # 示例用法
# if __name__ == "__main__":
#     # 用户可以替换为实际的txt文件路径和输出文件路径
#     txt_file_path = r"C:\Users\zhangbon\Desktop\temp.txt"
#     output_file_path = r"C:\Users\zhangbon\Desktop\output.txt"
    
#     result = read_and_join_lines(txt_file_path)
#     print("连接结果：", result)
    
#     # 将结果写入输出文件
#     try:
#         with open(output_file_path, 'w', encoding='utf-8') as f:
#             f.write(result)
#         print(f"结果已成功写入到 {output_file_path}")
#     except Exception as e:
#         print(f"写入文件时发生错误：{str(e)}")






with open(r"C:\Users\zhangbon\Desktop\temp.txt", encoding='utf-8') as f: 
    # 读取每一行文本
    lines = f.readlines()
    # 提取每行中的省份信息，假设格式为：省份名称+门店销售金额预测数据.xlsx
    province_list = []
    pattern = r'^处理完成：(.+?)门店销售金额预测数据\.xlsx'
    for line in lines:
        line = line.strip()
        match = re.search(pattern, line)
        if match:
            province_list.append(match.group(1))
print(province_list)
