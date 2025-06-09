"""
数据核对工具GUI应用
基于PyQt6的图形界面实现
"""

import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                            QLineEdit, QTextEdit, QProgressBar, QMessageBox,
                            QListWidget, QListWidgetItem)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
from datetime import datetime

# 导入原有的核对逻辑
from 终版 import read_excel_safely, compare_datasets, generate_detailed_report

class ComparisonWorker(QThread):
    """后台处理线程"""
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, source_path, target_paths, key_columns, output_dir):
        super().__init__()
        self.source_path = source_path
        self.target_paths = target_paths
        self.key_columns = [col.strip() for col in key_columns.replace('，', ',').split(',')]
        self.output_dir = output_dir

    def run(self):
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            
            # 读取源文件
            self.progress_signal.emit(f"正在加载源文件: {os.path.basename(self.source_path)}")
            source_df = read_excel_safely(self.source_path, self.key_columns)
            self.progress_signal.emit(f"源文件加载成功，记录数: {len(source_df)}")

            # 处理每个目标文件
            for target_path in self.target_paths:
                start_time = datetime.now()
                target_name = os.path.basename(target_path)
                self.progress_signal.emit(f"正在处理目标文件: {target_name}")

                try:
                    # 读取目标文件
                    target_df = read_excel_safely(target_path, self.key_columns)
                    self.progress_signal.emit(f"目标文件加载成功，记录数: {len(target_df)}")

                    # 执行比对
                    result_df = compare_datasets(source_df, target_df, self.key_columns)

                    # 生成报告
                    report_name = f"比对报告_{os.path.splitext(os.path.basename(self.source_path))[0]}_vs_{os.path.splitext(target_name)[0]}.xlsx"
                    output_path = os.path.join(self.output_dir, report_name)
                    generate_detailed_report(result_df, output_path)

                    # 统计信息
                    duration = (datetime.now() - start_time).total_seconds()
                    stats = f"""
完成比对 ({duration:.2f}秒)
差异统计：
- 仅存在于源文件: {len(result_df[result_df['主键状态'] == '仅存在于源文件'])}
- 仅存在于目标文件: {len(result_df[result_df['主键状态'] == '仅存在于目标文件'])}
- 存在数据差异: {len(result_df[result_df['主键状态'].str.contains('差异')])}
报告已保存至: {output_path}
"""
                    self.progress_signal.emit(stats)

                except Exception as e:
                    self.progress_signal.emit(f"处理失败: {str(e)}")

            self.finished_signal.emit(True, "处理完成")
        except Exception as e:
            self.finished_signal.emit(False, f"发生错误: {str(e)}")

class DataComparisonApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.source_columns = []  # 存储源文件的列名
        self.initUI()

    def initUI(self):
        self.setWindowTitle('数据核对工具')
        self.setGeometry(100, 100, 1000, 700)

        # 主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()

        # 源文件选择
        source_layout = QHBoxLayout()
        self.source_path = QLineEdit()
        source_btn = QPushButton('选择源文件')
        source_btn.clicked.connect(lambda: self.select_file(self.source_path))
        source_layout.addWidget(QLabel('源文件:'))
        source_layout.addWidget(self.source_path)
        source_layout.addWidget(source_btn)
        layout.addLayout(source_layout)

        # 目标文件列表
        target_label = QLabel('目标文件列表:')
        layout.addWidget(target_label)
        
        # 目标文件管理区域
        target_layout = QHBoxLayout()
        
        # 左侧：文件列表
        self.target_list = QListWidget()
        self.target_list.setMinimumHeight(150)
        
        # 右侧：按钮组
        button_layout = QVBoxLayout()
        add_target_btn = QPushButton('添加目标文件')
        remove_target_btn = QPushButton('删除选中文件')
        clear_target_btn = QPushButton('清空列表')
        
        add_target_btn.clicked.connect(self.add_target_file)
        remove_target_btn.clicked.connect(self.remove_target_file)
        clear_target_btn.clicked.connect(self.clear_target_list)
        
        button_layout.addWidget(add_target_btn)
        button_layout.addWidget(remove_target_btn)
        button_layout.addWidget(clear_target_btn)
        button_layout.addStretch()
        
        target_layout.addWidget(self.target_list, stretch=7)
        target_layout.addLayout(button_layout, stretch=3)
        layout.addLayout(target_layout)

        # 主键设置区域
        key_section = QVBoxLayout()
        
        # 主键输入行
        key_input_layout = QHBoxLayout()
        self.key_columns = QLineEdit()
        self.key_columns.setPlaceholderText('示例: 编号,姓名')
        preview_btn = QPushButton('预览列名')
        preview_btn.clicked.connect(self.preview_columns)
        key_input_layout.addWidget(QLabel('主键列(用英文逗号分隔):'))
        key_input_layout.addWidget(self.key_columns)
        key_input_layout.addWidget(preview_btn)
        
        # 列名预览区域
        self.columns_preview = QLineEdit()
        self.columns_preview.setReadOnly(True)
        self.columns_preview.setPlaceholderText('在这里显示源文件的所有列名...')
        
        key_section.addLayout(key_input_layout)
        key_section.addWidget(QLabel('可用列名(可直接复制):'))
        key_section.addWidget(self.columns_preview)
        layout.addLayout(key_section)

        # 输出目录选择
        output_layout = QHBoxLayout()
        self.output_dir = QLineEdit()
        output_btn = QPushButton('选择输出目录')
        output_btn.clicked.connect(lambda: self.select_directory(self.output_dir))
        output_layout.addWidget(QLabel('输出目录:'))
        output_layout.addWidget(self.output_dir)
        output_layout.addWidget(output_btn)
        layout.addLayout(output_layout)

        # 开始按钮
        self.start_btn = QPushButton('开始比对')
        self.start_btn.clicked.connect(self.start_comparison)
        layout.addWidget(self.start_btn)

        # 进度显示
        self.progress_text = QTextEdit()
        self.progress_text.setReadOnly(True)
        layout.addWidget(self.progress_text)

        main_widget.setLayout(layout)

    def add_target_file(self):
        """添加目标文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            '选择目标文件', 
            '', 
            'Excel文件 (*.xlsx *.xls)'
        )
        if file_path:
            item = QListWidgetItem(file_path)
            self.target_list.addItem(item)

    def remove_target_file(self):
        """删除选中的目标文件"""
        current_item = self.target_list.currentItem()
        if current_item:
            self.target_list.takeItem(self.target_list.row(current_item))

    def clear_target_list(self):
        """清空目标文件列表"""
        self.target_list.clear()

    def select_file(self, line_edit):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            '选择文件', 
            '', 
            'Excel文件 (*.xlsx *.xls)'
        )
        if file_path:
            line_edit.setText(file_path)
            # 当选择源文件后自动预览列名
            if line_edit == self.source_path:
                self.preview_columns()

    def select_directory(self, line_edit):
        directory = QFileDialog.getExistingDirectory(self, '选择目录')
        if directory:
            line_edit.setText(directory)

    def preview_columns(self):
        """预览源文件的列名"""
        if not self.source_path.text():
            QMessageBox.warning(self, '警告', '请先选择源文件')
            return
            
        try:
            # 读取前几行数据以获取列名
            df = pd.read_excel(self.source_path.text(), nrows=0)
            self.source_columns = df.columns.tolist()
            
            # 使用英文逗号横向展示列名
            self.columns_preview.setText(", ".join(self.source_columns))
            
        except Exception as e:
            QMessageBox.critical(self, '错误', f'读取列名失败: {str(e)}')

    def get_target_files(self):
        """获取所有目标文件路径"""
        return [
            self.target_list.item(i).text() 
            for i in range(self.target_list.count())
        ]

    def start_comparison(self):
        # 验证输入
        if not self.source_path.text() or \
           self.target_list.count() == 0 or \
           not self.key_columns.text() or \
           not self.output_dir.text():
            QMessageBox.warning(self, '警告', '请填写所有必要信息')
            return

        # 验证主键列是否存在
        input_keys = [k.strip() for k in self.key_columns.text().replace('，', ',').split(',')]
        if not all(key in self.source_columns for key in input_keys):
            invalid_keys = [key for key in input_keys if key not in self.source_columns]
            QMessageBox.warning(self, '警告', f'以下主键列在源文件中不存在：\n{", ".join(invalid_keys)}')
            return

        # 禁用开始按钮
        self.start_btn.setEnabled(False)
        self.progress_text.clear()

        # 创建并启动工作线程
        self.worker = ComparisonWorker(
            self.source_path.text(),
            self.get_target_files(),
            self.key_columns.text(),
            self.output_dir.text()
        )
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.comparison_finished)
        self.worker.start()

    def update_progress(self, message):
        self.progress_text.append(message)

    def comparison_finished(self, success, message):
        self.start_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, '完成', message)
        else:
            QMessageBox.critical(self, '错误', message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DataComparisonApp()
    ex.show()
    sys.exit(app.exec()) 