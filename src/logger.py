import datetime
import logging
import warnings
import os
from os.path import join

from pywinauto import actionlogger


class LogFilter(logging.Filter):
    def filter(self, record):
        message = record.getMessage()
        return 'WARNING! Cannot retrieve text length for handle' not in message


class PywinautoLoggerFilter(logging.Filter):
    def filter(self, record):
        return False


def setup_logger() -> None:
    root_folder = r'C:\Users\robot.ad\Desktop\sverka\logs'
    os.makedirs(root_folder, exist_ok=True)

    actionlogger.enable()
    actionlogger.ActionLogger.logger.propagate = True
    actionlogger.ActionLogger.logger.removeHandler(actionlogger.ActionLogger.logger.handlers[0])
    actionlogger.ActionLogger.logger.addFilter(LogFilter())

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter('%(asctime).19s %(levelname)s %(name)s %(filename)s %(funcName)s : %(message)s')

    today = datetime.date.today()
    year_month_folder = join(root_folder, today.strftime('%Y/%B'))
    os.makedirs(year_month_folder, exist_ok=True)

    file_handler = logging.FileHandler(
        join(year_month_folder, f'{today.strftime("%d.%m.%y")}.log'),
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.DEBUG)
    stream_handler.setFormatter(formatter)

    httpcore_logger = logging.getLogger('httpcore')
    httpcore_logger.setLevel(logging.INFO)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    # UserWarning: 32-bit application should be automated using 32-bit Python
    warnings.simplefilter(action='ignore', category=UserWarning)
