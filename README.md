# 发票识别工具

这是一个基于百度OCR API的发票识别工具，可以批量处理图片和PDF格式的发票，并将识别结果导出到Excel文件。

## 功能特点

- 支持JPG、PNG、JPEG和PDF格式的发票文件
- 使用百度OCR API进行高精度识别
- 简单易用的图形界面
- 实时显示处理进度
- 自动导出Excel结果文件

## 使用说明

1. 运行程序后，输入百度OCR API密钥
2. 点击"浏览"按钮选择需要处理的文件夹
3. 点击"开始识别"按钮开始处理
4. 等待处理完成，结果将自动保存为Excel文件

## 开发环境

- Python 3.8+
- PyQt5
- baidu-aip
- pdf2image
- openpyxl
- Pillow
- PyInstaller

## 安装依赖

```bash
pip install -r requirements.txt
```

## 打包说明

使用PyInstaller打包为独立可执行文件：

```bash
pyinstaller invoice_ocr.spec
```

## 注意事项

1. 使用前需要申请百度OCR API密钥
2. 确保文件夹中只包含支持的文件格式
3. 处理PDF文件时可能需要较长时间