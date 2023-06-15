import os
from os.path import join
import time
from contextlib import contextmanager
from datetime import date
from typing import List
import logger
import logging
from telegram import send_message
import platform


import win32com.client as win32

logger.setup_logger()

logging.info('Server started.')
logging.info('Logging started.')

PROJECT_FOLDER = r'C:\Users\robot.ad\Desktop\sverka-zp'
CHECK_INTERVAL: int = 60
RECIPIENTS: List[str] = ['robot.ad']
SUBJECT: str = 'test'
REPLIES_FILE: str = join(PROJECT_FOLDER, 'replied_emails.txt')
REPLY_MESSAGE: str = 'Добрый день, {}\n\n' \
                     'Ответ от робота\n\nСообщение сгенерировано автоматически.'
LACK_OF_ATTACHMENT_REPLY: str = 'Добрый день, {}\n\n' \
                                'Отсутствует вложенный файл.\n' \
                                'Пожалуйста приложите файл и отправьте новое отдельное письмо.' \
                                '\nПросьба не отвечать на это письмо.\n\n' \
                                'Сообщение сгенерировано автоматически.'

logging.info('Configuration loaded.')


@contextmanager
def dispatch(application: str) -> None:
    app = win32.Dispatch(application)
    namespace = app.GetNamespace('MAPI')
    try:
        yield namespace
    finally:
        app.Quit()


def save_reply(message_id: str) -> None:
    with open(REPLIES_FILE, 'a') as file:
        file.write(message_id + '\n')


def get_replied_messages() -> List[str]:
    with open(file=REPLIES_FILE, mode='r', encoding='utf-8') as file:
        replied_emails = file.read().splitlines()
    return replied_emails


def get_messages(outlook: win32.CDispatch) -> List[win32.CDispatch]:
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    if messages.Count == 0:
        return []
    messages.Sort('[ReceivedTime]', True)
    return [message for message in messages
            if hasattr(message, 'ReceivedTime') and
            message.ReceivedTime.date() == date.today() and
            message.Subject == SUBJECT and
            'RE:' not in message.Subject and
            message.SenderName in RECIPIENTS and
            message.EntryID not in get_replied_messages()]


def attachments_present(message: win32.CDispatch) -> bool:
    attachments = message.Attachments
    if attachments.Count == 0:
        return False
    return True


def run() -> None:
    if not os.path.exists(REPLIES_FILE):
        logging.info('File for replied emails not found. Creating replied emails file.')
        with open(REPLIES_FILE, 'w'):
            pass

    with dispatch('Outlook.Application') as outlook_namespace:
        logging.info('Outlook application started.')
        while True:
            try:
                logging.info('Checking inbox for new messages.')
                messages = get_messages(outlook_namespace)
                if not messages:
                    logging.info(f'No new messages. Waiting {CHECK_INTERVAL} seconds before checking for new messages.')
                    time.sleep(CHECK_INTERVAL)
                    continue
                logging.info(f'Found {len(messages)} new messages.')
                for message in messages:
                    if attachments_present(message):
                        logging.info(f'Sending succesful reply to {message.SenderName}.')
                        reply = message.Reply()
                        reply.Body = REPLY_MESSAGE.format(message.SenderName)
                        reply.Send()
                        logging.info(f'Succesful reply sent to {message.SenderName}.')
                    else:
                        logging.info(f'Sending reply of lack of attachment to {message.SenderName}.')
                        reply = message.Reply()
                        reply.Body = LACK_OF_ATTACHMENT_REPLY.format(message.SenderName)
                        logging.info(f'Reply of lack of attachment sent to {message.SenderName}.')
                        reply.Send()
                    save_reply(message_id=message.EntryID)
                    logging.info(f'Saved message id "{message.EntryID}" to replied emails file.')
                logging.info(f'All replies sent. Waiting {CHECK_INTERVAL} seconds before checking for new messages.')
                time.sleep(CHECK_INTERVAL)
            except Exception as error:
                logging.exception(f'An error occurred: {error}')
                send_message(f'Error occurred on {platform.node()}\nProcess: "Сверка зарплатной ведомости"\nError:\n{error}')
                raise error
            except KeyboardInterrupt as error:
                logging.exception(f'Keyboard interrupt. Terminating server.\n{error}')
                raise error


#
# if __name__ == '__main__':
#     while True:
#         check_outlook_inbox()
#         time.sleep(60)

# try:
#     data = get_input_data().data
#     robot = Robot(**data)
#     robot.run()
# except IndexError:
#     today = dt.now().strftime('%d.%m.%Y')
#     email = Email(
#         email_info=EmailInfo(email_list=ast.literal_eval({{email_list}})),
#         subject=f'Отсутствует файл для проверки {today}',
#         body=f'Отсутствует файл для проверки в \\\\dbu00234\\c$\\Temp\\Сверка\n\n'
#                 f'Пожалуйста добавьте его перед запуском робота.\n\nСообщение сгенерировано автоматически. '
#                 f'Просьба не отвечать.',
#     )
#     email.run()
