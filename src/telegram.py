import os
import dotenv
import httpx

dotenv.load_dotenv()
BOT_TOKEN = os.getenv('BOT_TOKEN')
CHAT_ID = os.getenv('CHAT_ID')


def send_message(message: str):
    return httpx.post(
        url=f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage',
        params={'chat_id': CHAT_ID, 'parse_mode': 'Markdown'},
        json={'text': message}
    )


if __name__ == '__main__':
    send_message(message='test')
