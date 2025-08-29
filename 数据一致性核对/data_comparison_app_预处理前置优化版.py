import sys
import os
import traceback
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                            QWidget, QPushButton, QLabel, QFileDialog, QTextEdit, 
                            QProgressBar, QTableWidget, QTableWidgetItem, QSplitter,
                            QGroupBox, QLineEdit, QMessageBox, QCheckBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtGui import QFont
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 导入预处理前置优化版的核心函数
from 一致性核对_预处理前置优化版 import DataPreprocessor, compare_preprocessed_datasets, generate_detailed_report_optimized

class SignalEmitter(QObject):
    """信号发射器"""
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

class ComparisonWorker(QThread):
    """后台比对工作线程"""
    
    def __init__(self, source_path, target_paths, key_columns, output_dir, use_cache=True):
        super().__init__()
        self.source_path = source_path
        self.target_paths = target_paths
        self.key_columns = key_columns
        self.output_dir = output_dir
        self.use_cache = use_cache
        self.signals = SignalEmitter()
        self.preprocessor = DataPreprocessor() if use_cache else None
        self._is_running = True
    
    def run(self):
        try:
            total_files = len(self.target_paths)
            self.signals.message.emit(f"开始预处理 {total_files + 1} 个文件...")
            
            # 预处理源数据
            self.signals.message.emit("预处理源数据...")
            source_df = self.preprocessor.preprocess_and_cache(self.source_path, self.key_columns)
            self.signals.progress.emit(10)
            
            # 预处理所有目标数据
            target_dfs = []
            for i, target_path in enumerate(self.target_paths):
                if not self._is_running:
                    break
                
                self.signals.message.emit(f"预处理目标文件 {i+1}/{total_files}: {os.path.basename(target_path)}")
                try:
                    target_df = self.preprocessor.preprocess_and_cache(target_path, self.key_columns)
                    target_dfs.append((target_path, target_df))
                    progress = 10 + int((i + 1) / total_files * 30)
                    self.signals.progress.emit(progress)
                except Exception as e:
                    self.signals.message.emit(f"⚠️ 预处理失败: {os.path.basename(target_path)} - {str(e)}")
            
            if not self._is_running:
                return
            
            # 开始比对
            self.signals.message.emit("开始数据比对...")
            
            for i, (target_path, target_df) in enumerate(target_dfs):
                if not self._is_running:
                    break
                
                target_name = os.path.basename(target_path)
                self.signals.message.emit(f"比对: {target_name}")
                
                try:
                    # 使用预处理后的数据进行比对
                    result_df = compare_preprocessed_datasets(source_df, target_df)
                    
                    # 生成报告
                    report_name = f"比对报告_{os.path.splitext(os.path.basename(self.source_path))[0]}_vs_{os.path.splitext(target_name)[0]}.xlsx"
                    output_path = os.path.join(self.output_dir, report_name)
                    generate_detailed_report_optimized(result_df, output_path)
                    
                    progress = 40 + int((i + 1) / len(target_dfs) * 60)
                    self.signals.progress.emit(progress)
                    
                    self.signals.message.emit(f"✅ 完成: {target_name}")
                    
                except Exception as e:
                    self.signals.message.emit(f"❌ 比对失败: {target_name} - {str(e)}")
            
            if self._is_running:
                self.signals.progress.emit(100)
                self.signals.message.emit("🎉 所有比对任务完成！")
                self.signals.finished.emit()
                
        except Exception as e:
            error_msg = f"严重错误: {str(e)}\n{traceback.format_exc()}"
            self.signals.error.emit(error_msg)
    
    def stop(self):
        """停止工作线程"""
        self._is_running = False

class DataComparisonApp(QMainWindow):
    """主应用程序窗口"""
    
    def __init__(self):
        super().__init__()
        self.source_file = ""
        self.target_files = []
        self.key_columns = []
        self.output_dir = ""
        self.worker = None
        
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("数据核对工具 - 预处理前置优化版")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)
        
        # 创建分割器
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 左侧面板 - 控制面板
        left_panel = self.create_control_panel()
        splitter.addWidget(left_panel)
        
        # 右侧面板 - 日志和预览
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)
        
        # 设置分割器比例
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        
        main_layout.addWidget(splitter)
    
    def create_control_panel(self):
        """创建控制面板"""
        panel = QGroupBox("控制面板")
        layout = QVBoxLayout()
        
        # 源文件选择
        source_group = QGroupBox("源文件")
        source_layout = QVBoxLayout()
        
        self.source_label = QLabel("未选择源文件")
        self.source_label.setWordWrap(True)
        
        source_btn_layout = QHBoxLayout()
        source_btn = QPushButton("选择源文件")
        source_btn.clicked.connect(self.select_source_file)
        source_btn_layout.addWidget(source_btn)
        
        source_layout.addWidget(self.source_label)
        source_layout.addLayout(source_btn_layout)
        source_group.setLayout(source_layout)
        
        # 目标文件选择
        target_group = QGroupBox("目标文件")
        target_layout = QVBoxLayout()
        
        self.target_table = QTableWidget()
        self.target_table.setColumnCount(2)
        self.target_table.setHorizontalHeaderLabels(["文件名", "状态"])
        self.target_table.setMaximumHeight(150)
        
        target_btn_layout = QHBoxLayout()
        add_target_btn = QPushButton("添加目标文件")
        add_target_btn.clicked.connect(self.add_target_files)
        clear_target_btn = QPushButton("清空")
        clear_target_btn.clicked.connect(self.clear_target_files)
        
        target_btn_layout.addWidget(add_target_btn)
        target_btn_layout.addWidget(clear_target_btn)
        
        target_layout.addWidget(self.target_table)
        target_layout.addLayout(target_btn_layout)
        target_group.setLayout(target_layout)
        
        # 主键设置
        key_group = QGroupBox("主键设置")
        key_layout = QVBoxLayout()
        
        # 源文件列名显示
        source_columns_layout = QVBoxLayout()
        source_columns_layout.addWidget(QLabel("源文件列名:"))
        self.source_columns_label = QLineEdit("请先选择源文件")
        self.source_columns_label.setReadOnly(True)
        self.source_columns_label.setStyleSheet("background-color: #f0f0f0; padding: 5px; font-family: monospace; min-height: 30px;")
        source_columns_layout.addWidget(self.source_columns_label)
        
        # 目标文件列名显示
        target_columns_layout = QVBoxLayout()
        target_columns_layout.addWidget(QLabel("目标文件列名:"))
        self.target_columns_label = QLineEdit("请先选择目标文件")
        self.target_columns_label.setReadOnly(True)
        self.target_columns_label.setStyleSheet("background-color: #f0f0f0; padding: 5px; font-family: monospace; min-height: 30px;")
        target_columns_layout.addWidget(self.target_columns_label)
        
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("输入主键列名，用逗号分隔")
        self.key_input.textChanged.connect(self.normalize_key_input)
        
        key_layout.addLayout(source_columns_layout)
        key_layout.addLayout(target_columns_layout)
        key_layout.addWidget(QLabel("主键列名:"))
        key_layout.addWidget(self.key_input)
        key_group.setLayout(key_layout)
        
        # 输出目录
        output_group = QGroupBox("输出目录")
        output_layout = QVBoxLayout()
        
        self.output_label = QLabel("未选择输出目录")
        self.output_label.setWordWrap(True)
        output_btn = QPushButton("选择输出目录")
        output_btn.clicked.connect(self.select_output_dir)
        
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(output_btn)
        output_group.setLayout(output_layout)
        
        # 缓存选项
        cache_group = QGroupBox("缓存设置")
        cache_layout = QVBoxLayout()
        self.cache_checkbox = QCheckBox("启用预处理缓存")
        self.cache_checkbox.setChecked(True)
        cache_layout.addWidget(self.cache_checkbox)
        cache_group.setLayout(cache_layout)
        
        # 开始按钮
        self.start_btn = QPushButton("开始比对")
        self.start_btn.clicked.connect(self.start_comparison)
        self.start_btn.setEnabled(False)
        
        # 停止按钮
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_comparison)
        self.stop_btn.setEnabled(False)
        
        # 添加到布局
        layout.addWidget(source_group)
        layout.addWidget(target_group)
        layout.addWidget(key_group)
        layout.addWidget(output_group)
        layout.addWidget(cache_group)
        layout.addWidget(self.start_btn)
        layout.addWidget(self.stop_btn)
        layout.addStretch()
        
        panel.setLayout(layout)
        return panel
    
    def create_right_panel(self):
        """创建右侧面板"""
        panel = QWidget()
        layout = QVBoxLayout()
        
        # 进度条
        progress_group = QGroupBox("进度")
        progress_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        progress_group.setLayout(progress_layout)
        
        # 日志显示
        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(300)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        
        layout.addWidget(progress_group)
        layout.addWidget(log_group)
        layout.addStretch()
        
        panel.setLayout(layout)
        return panel
    
    def select_source_file(self):
        """选择源文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择源文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.source_file = file_path
            self.source_label.setText(os.path.basename(file_path))
            self.update_source_columns(file_path)
            self.check_ready()
    
    def add_target_files(self):
        """添加目标文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择目标文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        
        for file_path in file_paths:
            if file_path not in self.target_files:
                self.target_files.append(file_path)
                self.update_target_table()
                self.update_target_columns()
                self.check_ready()
    
    def clear_target_files(self):
        """清空目标文件"""
        self.target_files.clear()
        self.update_target_table()
        self.target_columns_label.setText("请先选择目标文件")
        self.check_ready()
    
    def update_source_columns(self, file_path):
        """更新源文件列名显示"""
        try:
            df = pd.read_excel(file_path, nrows=0)
            columns = df.columns.tolist()
            columns_text = ', '.join(columns)
            self.source_columns_label.setText(columns_text)
        except Exception as e:
            self.source_columns_label.setText(f"读取列名失败: {str(e)}")
    
    def update_target_columns(self):
        """更新目标文件列名显示"""
        if not self.target_files:
            self.target_columns_label.setText("请先选择目标文件")
            return
        
        if len(self.target_files) == 1:
            try:
                df = pd.read_excel(self.target_files[0], nrows=0)
                columns = df.columns.tolist()
                columns_text = ', '.join(columns)
                self.target_columns_label.setText(columns_text)
            except Exception as e:
                self.target_columns_label.setText(f"读取列名失败: {str(e)}")
        else:
            # 多个文件时显示第一个文件的列名
            try:
                df = pd.read_excel(self.target_files[0], nrows=0)
                columns = df.columns.tolist()
                columns_text = ', '.join(columns)
                file_count = len(self.target_files)
                self.target_columns_label.setText(f"文件1: {columns_text}\n(共{file_count}个文件)")
            except Exception as e:
                self.target_columns_label.setText(f"读取列名失败: {str(e)}")
    
    def update_target_table(self):
        """更新目标文件表格"""
        self.target_table.setRowCount(len(self.target_files))
        for row, file_path in enumerate(self.target_files):
            self.target_table.setItem(row, 0, QTableWidgetItem(os.path.basename(file_path)))
            self.target_table.setItem(row, 1, QTableWidgetItem("待处理"))
    
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir = dir_path
            self.output_label.setText(dir_path)
            self.check_ready()
    
    def check_ready(self):
        """检查是否可以开始比对"""
        ready = bool(
            self.source_file and 
            self.target_files and 
            self.output_dir and 
            self.key_input.text().strip()
        )
        self.start_btn.setEnabled(ready)
    
    def normalize_key_input(self, text):
        """自动将中文逗号转为英文逗号"""
        if '，' in text:
            normalized_text = text.replace('，', ',')
            self.key_input.setText(normalized_text)
    
    def start_comparison(self):
        """开始比对"""
        if not self.validate_inputs():
            return
        
        # 获取主键列（已自动转换中文逗号）
        key_text = self.key_input.text().strip()
        self.key_columns = [col.strip() for col in key_text.split(',')]
        
        # 更新UI状态
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        
        # 创建工作线程
        self.worker = ComparisonWorker(
            self.source_file, 
            self.target_files, 
            self.key_columns, 
            self.output_dir,
            self.cache_checkbox.isChecked()
        )
        
        # 连接信号
        self.worker.signals.progress.connect(self.update_progress)
        self.worker.signals.message.connect(self.append_log)
        self.worker.signals.finished.connect(self.comparison_finished)
        self.worker.signals.error.connect(self.show_error)
        
        # 启动线程
        self.worker.start()
    
    def validate_inputs(self):
        """验证输入"""
        if not self.source_file:
            QMessageBox.warning(self, "警告", "请选择源文件")
            return False
        
        if not self.target_files:
            QMessageBox.warning(self, "警告", "请选择至少一个目标文件")
            return False
        
        if not self.output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return False
        
        if not self.key_input.text().strip():
            QMessageBox.warning(self, "警告", "请输入主键列名")
            return False
        
        return True
    
    def stop_comparison(self):
        """停止比对"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.quit()
            self.worker.wait()
        
        self.comparison_finished()
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def append_log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def comparison_finished(self):
        """比对完成"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.worker = None
    
    def show_error(self, error_msg):
        """显示错误"""
        QMessageBox.critical(self, "错误", error_msg)
        self.comparison_finished()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataComparisonApp()
    window.show()
    sys.exit(app.exec())