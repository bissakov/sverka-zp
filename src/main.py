import logging
import platform
import traceback

from server import run
from src.telegram import send_picture


if __name__ == '__main__':
    try:
        run()
    except Exception as error:
        error_message = traceback.format_exc()
        logging.exception(error)
        send_picture(caption=f'Error occurred on {platform.node()}\n'
                             f'Process: "Сверка зарплатной ведомости"\nError:\n{error_message}')
        raise error
    except KeyboardInterrupt as error:
        logging.info('Server manually stopped via KeyboardInterrupt.')
        raise error
