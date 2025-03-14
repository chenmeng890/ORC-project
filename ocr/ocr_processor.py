from aip import AipOcr
from PIL import Image
from pdf2image import convert_from_path
import os
import logging
import tempfile

class OCRProcessor:
    def __init__(self, api_key):
        # 解析API密钥字符串
        keys = api_key.strip().split(',')
        if len(keys) != 3:
            raise ValueError('API密钥格式错误，请确保包含APP_ID,API_KEY,SECRET_KEY三个值')
        
        app_id, api_key, secret_key = keys
        self.client = AipOcr(app_id, api_key, secret_key)

    def process_image(self, image_path):
        try:
            logging.info(f'开始处理图片：{image_path}')
            with open(image_path, 'rb') as fp:
                image = fp.read()
            logging.info('正在调用百度OCR API...')
            result = self.client.vatInvoice(image)
            
            if 'error_code' in result:
                error_msg = f'百度API返回错误：错误码 {result["error_code"]}，错误信息：{result.get("error_msg", "未知错误")}'
                logging.error(error_msg)
                raise Exception(error_msg)
            
            logging.info(f'API调用成功，返回结果：{result}')
            return [result.get('words_result', {})] if 'words_result' in result else []
        except Exception as e:
            logging.error(f'处理图片 {image_path} 时出错：{str(e)}')
            raise e

    def process_pdf(self, pdf_path):
        try:
            # 设置转换参数以提高图像质量
            pages = convert_from_path(
                pdf_path,
                dpi=300,  # 提高DPI以获取更清晰的图像
                fmt='jpeg',
                grayscale=False,  # 保持彩色以避免信息丢失
                thread_count=2  # 使用多线程加速处理
            )
            
            if not pages:
                error_msg = f'PDF文件 {pdf_path} 转换失败：未能提取到页面'
                logging.error(error_msg)
                raise Exception(error_msg)
                
            results = []
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=True) as temp_file:
                for page_num, page in enumerate(pages, 1):
                    try:
                        # 优化图像保存质量
                        page.save(temp_file.name, 'JPEG', quality=95)
                        result = self.process_image(temp_file.name)
                        if result:
                            results.extend(result)
                        else:
                            logging.warning(f'PDF文件 {pdf_path} 第{page_num}页识别结果为空')
                    except Exception as e:
                        logging.error(f'处理PDF文件 {pdf_path} 第{page_num}页时出错：{str(e)}')
                        continue
                        
            if not results:
                logging.warning(f'PDF文件 {pdf_path} 未能成功识别任何内容')
            return results
        except Exception as e:
            error_msg = f'处理PDF文件 {pdf_path} 时出错：{str(e)}'
            logging.error(error_msg)
            return []

    def process_file(self, file_path):
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext in ['.jpg', '.jpeg', '.png']:
            return self.process_image(file_path)
        elif file_ext == '.pdf':
            return self.process_pdf(file_path)
        else:
            logging.warning(f'不支持的文件格式：{file_ext}')
            return []