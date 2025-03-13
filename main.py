import sys
import logging
from PyQt5.QtWidgets import QApplication
from gui.main_window import MainWindow

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('invoice_ocr.log')
        ]
    )

def main():
    try:
        setup_logging()
        logging.info('启动发票识别程序')
        
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        
        exit_code = app.exec_()
        logging.info('程序正常退出')
        sys.exit(exit_code)
        
    except Exception as e:
        logging.error(f'程序发生异常：{str(e)}')
        sys.exit(1)

if __name__ == '__main__':
    main()