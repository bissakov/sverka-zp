import io
import logging
from time import sleep
from urllib.parse import urljoin

import requests
from PIL import Image, ImageGrab

from config import BOT_TOKEN, CHAT_ID, SESSION


class TelegramBot:
    def __init__(self, token: str = BOT_TOKEN, chat_id: str = CHAT_ID, session: requests.Session = SESSION) -> None:
        self.session = session
        self.api_url = f'https://api.telegram.org/bot{token}/'
        self.send_data = {'chat_id': chat_id}

    def send_message(self, message: str) -> bool:
        self.send_data['text'] = message
        url = urljoin(self.api_url, 'sendMessage')
        response = self.session.post(url, data=self.send_data)
        logging.info(f'response: {response}')
        return response.status_code == 200

    def send_document(self, document_path: str, caption: str | None = None) -> bool:
        document_file = {'document': open(file=document_path, mode='rb')}
        if caption:
            self.send_data['caption'] = caption
        url = urljoin(self.api_url, 'sendDocument')
        response = self.session.post(url, data=self.send_data, files=document_file)
        logging.info(f'response: {response}')
        return response.status_code == 200

    def send_picture(self, image: Image.Image = None, caption: str | None = None) -> bool:
        if not image:
            image = ImageGrab.grab()
        image_stream = io.BytesIO()
        image.save(image_stream, format='JPEG', quality=70)

        image_stream.seek(0)
        image_file = {'photo': ('image.png', image_stream)}

        if caption:
            self.send_data['caption'] = caption
        url = urljoin(self.api_url, 'sendPhoto')
        response = self.session.post(url, data=self.send_data, files=image_file)
        return response.status_code == 200


def send_message(message: str) -> None:
    bot = TelegramBot()
    bot.send_message(message=message)


def send_document(document_path: str, caption: str | None = None) -> None:
    bot = TelegramBot()
    bot.send_document(document_path=document_path, caption=caption)


def send_picture(image: Image.Image = None, caption: str | None = None) -> None:
    bot = TelegramBot()
    bot.send_picture(image=image, caption=caption)


def foo():
    send_message('test')
    send_document(document_path=r"C:\Users\robot.ad\Desktop\sverka-zp\excel_reports\Prov_JYL_2023.xlsx", caption='adsadads')
    send_picture(caption='sadsadasdasdas')

    pass


if __name__ == '__main__':
    foo()
