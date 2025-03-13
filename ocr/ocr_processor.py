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
            with open(image_path, 'rb') as fp:
                image = fp.read()
            result = self.client.vatInvoice(image)
            return [result.get('words_result', {})] if 'words_result' in result else []
        except Exception as e:
            logging.error(f'处理图片 {image_path} 时出错：{str(e)}')
            return []

    def process_pdf(self, pdf_path):
        try:
            pages = convert_from_path(pdf_path)
            results = []
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=True) as temp_file:
                for page in pages:
                    # 将PDF页面转换为图片
                    page.save(temp_file.name, 'JPEG')
                    result = self.process_image(temp_file.name)
                    results.extend(result)
            return results
        except Exception as e:
            logging.error(f'处理PDF {pdf_path} 时出错：{str(e)}')
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