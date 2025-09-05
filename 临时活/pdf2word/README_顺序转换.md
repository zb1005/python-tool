# PDF转Word顺序转换工具

## 功能特性

- ✅ 按照指定顺序列表转换PDF文件
- ✅ 自动合并所有转换后的Word文档为一个文件
- ✅ 支持命令行参数配置
- ✅ 详细的日志记录
- ✅ 创建示例顺序列表文件

## 安装依赖

```bash
pip install pdf2docx python-docx
```

## 使用方法

### 1. 创建顺序列表文件

创建一个文本文件（如 `order_list.txt`），每行一个PDF文件路径，按需要的顺序排列：

```txt
# PDF文件顺序列表
# 每行一个PDF文件路径，按顺序排列
# 注释以#开头

D:\文档\第一章.pdf
D:\文档\第二章.pdf
D:\文档\第三章.pdf
```

### 2. 运行顺序转换

```bash
python pdf2word.py --order-file order_list.txt --output-dir output_word
```

### 3. 创建示例顺序列表文件

```bash
python pdf2word.py --create-sample
```

## 命令行参数

- `--order-file` 或 `-o`: 顺序列表文件路径（必需）
- `--output-dir` 或 `-d`: 输出目录（必需）
- `--create-sample`: 创建示例顺序列表文件

## 输出结果

- 每个PDF文件会被转换为单独的Word文档
- 所有转换后的Word文档会被合并到 `merged_document.docx`
- 转换日志保存在 `pdf_to_word_conversion.log`

## 示例

```bash
# 示例1：使用顺序列表转换
python pdf2word.py -o my_order.txt -d output_files

# 示例2：创建示例顺序列表
python pdf2word.py --create-sample
```

## 注意事项

1. 确保顺序列表文件中的PDF路径正确
2. 输出目录需要有写入权限
3. 合并后的文档会按顺序列表中的顺序排列
4. 每个PDF转换后的文档会单独保存，同时也会合并为一个总文档