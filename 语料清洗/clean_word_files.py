import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def clean_word_files(folder_path):
    # 遍历文件夹中的所有文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            file_path = os.path.join(folder_path, filename)
            try:
                # 打开Word文档
                doc = Document(file_path)
                
                # 处理每个段落
                for para in doc.paragraphs:
                    # 保存原始格式
                    original_style = para.style
                    original_align = para.alignment
                    original_font_size = None
                    if para.runs:
                        original_font_size = para.runs[0].font.size
                    
                    # 清洗文本: 删除换行符和.0
                    cleaned_text = para.text.replace(r'\n', '').replace('.0', '')
                    cleaned_text = cleaned_text.replace(',"Unnamed: 9":""', '')
                    cleaned_text = cleaned_text.replace(',"Unnamed: 10":""', '')
                    cleaned_text = cleaned_text.replace(',"Unnamed: 11":""', '')
                    cleaned_text = cleaned_text.replace(',"Unnamed: 12":""', '')    
                    # 清空段落并添加清洗后的文本
                    para.text = cleaned_text
                    
                    # 恢复原始格式
                    para.style = original_style
                    para.alignment = original_align
                    if original_font_size and para.runs:
                        para.runs[0].font.size = original_font_size
                
                # 保存修改后的文档
                doc.save(file_path)
                print(f'已处理: {filename}')
            except Exception as e:
                print(f'处理{filename}时出错: {str(e)}')

if __name__ == '__main__':
    # 指定要处理的文件夹路径
    target_folder = os.path.dirname(os.path.abspath(__file__))  # 当前脚本所在文件夹
    # 或者手动指定文件夹路径
    target_folder = r'C:\Users\zhangbon\Desktop\2025_06_12_AI知识库\AI知识库——20250703\数据消费—20250627'
    
    if os.path.exists(target_folder) and os.path.isdir(target_folder):
        clean_word_files(target_folder)
        print('所有Word文档处理完成!')
    else:
        print(f'错误: 文件夹不存在或不是有效的目录 - {target_folder}')