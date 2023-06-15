import datetime
import logging
from os import makedirs
from os.path import join

from pywinauto import actionlogger


class LogFilter(logging.Filter):
    def filter(self, record):
        return 'Cannot retrieve text length for handle' not in record.getMessage()


def setup_logger() -> None:
    root_folder = r'C:\Users\robot.ad\Desktop\sverka-zp\logs'
    makedirs(root_folder, exist_ok=True)
    actionlogger.enable()
    logger = logging.getLogger()
    # logger.setLevel(logging.INFO)
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter('%(asctime).19s %(levelname)s %(name)s %(threadName)s : %(message)s')

    file_handler = logging.FileHandler(
        join(root_folder, f'{datetime.date.today().strftime("%d.%m.%y")}.log'),
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.DEBUG)
    stream_handler.setFormatter(formatter)

    pywinauto_logger = logging.getLogger('pywinauto')
    pywinauto_logger.addFilter(LogFilter())

    httpx_logger = logging.getLogger('httpx')
    httpx_logger.setLevel(logging.INFO)

    httpcore_logger = logging.getLogger('httpcore')
    httpcore_logger.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
