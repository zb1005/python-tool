import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                            QLineEdit, QTextEdit, QProgressBar, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import pandas as pd

class SimpleComparisonWorker(QThread):
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, source_path, target_path, key_columns, output_path):
        super().__init__()
        self.source_path = source_path
        self.target_path = target_path
        self.key_columns = [col.strip() for col in key_columns.replace('，', ',').split(',')]
        self.output_path = output_path

    def run(self):
        try:
            self.progress_signal.emit(f"正在加载源文件: {os.path.basename(self.source_path)}")
            source_df = pd.read_excel(self.source_path, dtype=str, keep_default_na=False)
            source_df.columns = source_df.columns.str.strip()
            # 去除所有字符串数据的前后空格
            source_df = source_df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
            # 将所有空值填充为空格
            source_df = source_df.fillna(' ')
            
            self.progress_signal.emit(f"正在加载目标文件: {os.path.basename(self.target_path)}")
            target_df = pd.read_excel(self.target_path, dtype=str, keep_default_na=False)
            target_df.columns = target_df.columns.str.strip()
            # 去除所有字符串数据的前后空格
            target_df = target_df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
            # 将所有空值填充为空格
            target_df = target_df.fillna(' ')
            
            self.progress_signal.emit("开始比对数据...")
            result_df = self.compare_datasets(source_df, target_df, self.key_columns)
            
            result_df.to_excel(self.output_path, index=False)
            self.finished_signal.emit(True, f"比对完成，结果已保存至: {self.output_path}")
        except Exception as e:
            self.finished_signal.emit(False, f"处理失败: {str(e)}")

    def compare_datasets(self, source_df, target_df, key_columns):
        """核心比对函数（字典优化版）"""
        # 生成主键并检查重复
        source_df['_composite_key'] = source_df.apply(
            lambda x: '_'.join([
                str(x[col]).strip() if pd.notna(x[col]) and str(x[col]).strip() != ''
                else '<空值>'
                for col in key_columns
            ]), axis=1)
        target_df['_composite_key'] = target_df.apply(
            lambda x: '_'.join([
                str(x[col]).strip() if pd.notna(x[col]) and str(x[col]).strip() != ''
                else '<空值>'
                for col in key_columns
            ]), axis=1)
        
        # 检查主键唯一性
        if source_df['_composite_key'].duplicated().any():
            dup_keys = source_df[source_df['_composite_key'].duplicated()]['_composite_key'].unique()
            raise ValueError(f"源数据中存在重复主键: {','.join(dup_keys[:3])}...")
        if target_df['_composite_key'].duplicated().any():
            dup_keys = target_df[target_df['_composite_key'].duplicated()]['_composite_key'].unique()
            raise ValueError(f"目标数据中存在重复主键: {','.join(dup_keys[:3])}...")
        
        # 转换为字典提高查找效率
        source_dict = source_df.set_index('_composite_key').to_dict('index')
        target_dict = target_df.set_index('_composite_key').to_dict('index')
        
        # 获取主键集合和共同列
        source_keys = set(source_dict.keys())
        target_keys = set(target_dict.keys())
        common_columns = [col for col in source_df.columns if col in target_df.columns and col != '_composite_key']
        
        results = []

        # 处理仅存在于源数据的主键
        for key in source_keys - target_keys:
            results.append({
                '主键状态': '仅存在于源文件',
                '组合主键': key,
                **{f"源_{col}": source_dict[key][col] for col in common_columns},
                **{f"目标_{col}": "" for col in common_columns}
            })

        # 处理仅存在于目标数据的主键
        for key in target_keys - source_keys:
            results.append({
                '主键状态': '仅存在于目标文件',
                '组合主键': key,
                **{f"源_{col}": "" for col in common_columns},
                **{f"目标_{col}": target_dict[key][col] for col in common_columns}
            })

        # 处理共同主键的数据差异
        for key in source_keys & target_keys:
            src_row = source_dict[key]
            tgt_row = target_dict[key]

            diff_details = {}
            row_data = {'主键状态': '数据一致', '组合主键': key}

            for col in common_columns:
                src_val = src_row[col]
                tgt_val = tgt_row[col]
                
                if src_val != tgt_val:
                    diff_details[col] = {'源值': src_val, '目标值': tgt_val}
                
                row_data[f"源_{col}"] = src_val
                row_data[f"目标_{col}"] = tgt_val

            if diff_details:
                row_data['主键状态'] = f"发现{len(diff_details)}处差异"
                row_data['差异详情'] = str(diff_details)
                row_data['差异列名'] = ', '.join(diff_details.keys())

            results.append(row_data)

        return pd.DataFrame(results)

class SimpleComparisonApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('简化版数据比对工具')
        self.setGeometry(100, 100, 800, 600)

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

        # 目标文件选择
        target_layout = QHBoxLayout()
        self.target_path = QLineEdit()
        target_btn = QPushButton('选择目标文件')
        target_btn.clicked.connect(lambda: self.select_file(self.target_path))
        target_layout.addWidget(QLabel('目标文件:'))
        target_layout.addWidget(self.target_path)
        target_layout.addWidget(target_btn)
        layout.addLayout(target_layout)

        # 主键输入
        key_layout = QHBoxLayout()
        self.key_columns = QLineEdit()
        self.key_columns.setPlaceholderText('输入主键列，用逗号分隔')
        key_layout.addWidget(QLabel('主键列:'))
        key_layout.addWidget(self.key_columns)
        layout.addLayout(key_layout)

        # 输出文件选择
        output_layout = QHBoxLayout()
        self.output_path = QLineEdit()
        output_btn = QPushButton('选择输出文件')
        output_btn.clicked.connect(self.select_output_file)
        output_layout.addWidget(QLabel('输出文件:'))
        output_layout.addWidget(self.output_path)
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

    def select_file(self, line_edit):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择文件', '', 'Excel文件 (*.xlsx *.xls)')
        if file_path:
            line_edit.setText(file_path)

    def select_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(self, '保存结果', '', 'Excel文件 (*.xlsx)')
        if file_path:
            self.output_path.setText(file_path)

    def start_comparison(self):
        if not all([self.source_path.text(), self.target_path.text(), 
                   self.key_columns.text(), self.output_path.text()]):
            QMessageBox.warning(self, '警告', '请填写所有必要信息')
            return

        self.start_btn.setEnabled(False)
        self.progress_text.clear()

        self.worker = SimpleComparisonWorker(
            self.source_path.text(),
            self.target_path.text(),
            self.key_columns.text(),
            self.output_path.text()
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
    ex = SimpleComparisonApp()
    ex.show()
    sys.exit(app.exec())