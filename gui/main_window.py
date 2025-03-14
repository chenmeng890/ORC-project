from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLineEdit, QLabel, QFileDialog,
                             QProgressBar, QTextEdit, QSplitter)
from PyQt5.QtCore import Qt
from ocr.ocr_processor import OCRProcessor
from data.excel_processor import ExcelProcessor
import os
import json

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('发票识别工具')
        self.setMinimumSize(800, 600)
        self.init_ui()
        self.ocr_processor = None
        self.excel_processor = ExcelProcessor()
        self.excel_processor.set_callback(self.handle_excel_progress)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # API密钥输入区域
        api_layout = QVBoxLayout()
        api_layout.setSpacing(10)
        
        # APP_ID输入框
        app_id_layout = QHBoxLayout()
        app_id_label = QLabel('APP_ID：')
        self.app_id_input = QLineEdit()
        self.app_id_input.setText('118016480')
        app_id_layout.addWidget(app_id_label)
        app_id_layout.addWidget(self.app_id_input)
        api_layout.addLayout(app_id_layout)
        
        # API_KEY输入框
        api_key_layout = QHBoxLayout()
        api_key_label = QLabel('API_KEY：')
        self.api_key_input = QLineEdit()
        self.api_key_input.setText('umqnL6BlsB1mRtt6rNfjsfda')
        api_key_layout.addWidget(api_key_label)
        api_key_layout.addWidget(self.api_key_input)
        api_layout.addLayout(api_key_layout)
        
        # SECRET_KEY输入框
        secret_key_layout = QHBoxLayout()
        secret_key_label = QLabel('SECRET_KEY：')
        self.secret_key_input = QLineEdit()
        self.secret_key_input.setText('tcVw9RkZbEMZmBUm02cicsEuwLBjzlWo')
        secret_key_layout.addWidget(secret_key_label)
        secret_key_layout.addWidget(self.secret_key_input)
        api_layout.addLayout(secret_key_layout)
        
        layout.addLayout(api_layout)

        # 文件夹选择区域
        folder_layout = QHBoxLayout()
        self.folder_path_label = QLabel('选择文件夹：')
        self.folder_path_input = QLineEdit()
        self.folder_select_btn = QPushButton('浏览')
        self.folder_select_btn.clicked.connect(self.select_folder)
        self.open_folder_btn = QPushButton('打开文件夹')
        self.open_folder_btn.clicked.connect(self.open_selected_folder)
        folder_layout.addWidget(self.folder_path_label)
        folder_layout.addWidget(self.folder_path_input)
        folder_layout.addWidget(self.folder_select_btn)
        folder_layout.addWidget(self.open_folder_btn)
        layout.addLayout(folder_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # 创建一个分割器来容纳两个文本显示区域
        splitter = QSplitter(Qt.Vertical)
        layout.addWidget(splitter)

        # 日志显示区域
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        splitter.addWidget(self.log_display)

        # OCR原始数据显示区域
        self.ocr_data_display = QTextEdit()
        self.ocr_data_display.setReadOnly(True)
        splitter.addWidget(self.ocr_data_display)

        # 开始按钮
        self.start_btn = QPushButton('开始识别')
        self.start_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.start_btn)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if folder:
            self.folder_path_input.setText(folder)

    def start_processing(self):
        app_id = self.app_id_input.text()
        api_key = self.api_key_input.text()
        secret_key = self.secret_key_input.text()
        folder_path = self.folder_path_input.text()

        if not app_id or not api_key or not secret_key or not folder_path:
            self.log_display.append('错误：请输入所有API密钥和选择文件夹')
            return

        try:
            api_key_str = f'{app_id},{api_key},{secret_key}'
            self.ocr_processor = OCRProcessor(api_key_str)
            self.excel_processor.create_new_workbook()
            self.start_btn.setEnabled(False)
            self.process_folder(folder_path)
        except ValueError as e:
            self.log_display.append(f'错误：{str(e)}')
            return

    def handle_excel_progress(self, filename, status, message):
        if status == 'success':
            self.log_display.append(f'✅ {filename}: {message}')
        else:
            self.log_display.append(f'❌ {filename}: {message}')

    def process_folder(self, folder_path):
        supported_extensions = ['.jpg', '.jpeg', '.png', '.pdf']
        files = [f for f in os.listdir(folder_path) 
                if os.path.splitext(f)[1].lower() in supported_extensions]
        total_files = len(files)

        if total_files == 0:
            self.log_display.append('错误：所选文件夹中没有支持的文件格式')
            self.start_btn.setEnabled(True)
            return

        self.progress_bar.setMaximum(total_files)
        processed_files = 0

        for filename in files:
            file_path = os.path.join(folder_path, filename)
            self.log_display.append(f'正在处理：{filename}')
            
            try:
                ocr_result = self.ocr_processor.process_file(file_path)
                # 显示OCR原始数据
                self.ocr_data_display.append(f'文件：{filename}')
                self.ocr_data_display.append('OCR识别结果：')
                self.ocr_data_display.append(json.dumps(ocr_result, ensure_ascii=False, indent=2))
                self.ocr_data_display.append('\n' + '-'*50 + '\n')
                
                self.excel_processor.add_result(filename, ocr_result)
                processed_files += 1
                self.progress_bar.setValue(processed_files)
            except Exception as e:
                self.log_display.append(f'处理文件 {filename} 时出错：{str(e)}')
                continue

        # 保存Excel文件
        try:
            excel_path = os.path.join(folder_path, '发票识别结果.xlsx')
            self.excel_processor.save_workbook(excel_path)
            self.log_display.append(f'处理完成！结果已保存至：{excel_path}')
        except Exception as e:
            self.log_display.append(f'保存Excel文件时出错：{str(e)}')

        self.start_btn.setEnabled(True)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def log_message(self, message):
        self.log_display.append(message)

    def open_selected_folder(self):
        folder_path = self.folder_path_input.text()
        if folder_path and os.path.exists(folder_path):
            os.system(f'open "{folder_path}"')
        else:
            self.log_display.append('错误：请先选择有效的文件夹')