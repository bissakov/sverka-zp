import os
from os.path import join
import time
from contextlib import contextmanager
from datetime import date
from typing import Any, List
import logger
import logging
from telegram import send_message
import platform

from sqlalchemy import ColumnElement, create_engine, Column, String
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base

import win32com.client as win32


logger.setup_logger()

logging.info('Server started.')
logging.info('Logging started.')

Base = declarative_base()


class Reply(Base):
    __tablename__ = 'replied_emails'
    message_id = Column(String, primary_key=True)


db_root_folder = r'C:\Users\robot.ad\Desktop\sverka-zp\database'
os.makedirs(db_root_folder, exist_ok=True)
engine = create_engine(f'sqlite:///{db_root_folder}/replies.db')
Session = sessionmaker(bind=engine)
SESSION = Session()
Base.metadata.create_all(engine)
logging.info('SQLite session started.')


PROJECT_FOLDER = r'C:\Users\robot.ad\Desktop\sverka-zp'
CHECK_INTERVAL: int = 60
RECIPIENTS: List[str] = ['robot.ad']
SUBJECT: str = 'test'
REPLIES_FILE: str = join(PROJECT_FOLDER, r'replied_emails\replied_emails.txt')
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
    replied_email = Reply(message_id=message_id)
    SESSION.add(replied_email)
    SESSION.commit()


def get_replied_messages() -> list[ColumnElement[Any]]:
    replied_emails = SESSION.query(Reply.message_id).all()
    replied_emails = [email[0] for email in replied_emails]
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


def reply_to_message(message: win32.CDispatch, reply_message: str) -> None:
    reply = message.Reply()
    reply.Body = reply_message
    reply.Send()
    save_reply(message.EntryID)
    logging.info(f'Saved message id "{message.EntryID}" to replied emails file.')


def send_reply(message: win32.CDispatch) -> None:
    if attachments_present(message):
        logging.info(f'Sending succesful reply to {message.SenderName}.')
        reply_to_message(message, REPLY_MESSAGE.format(message.SenderName))
        logging.info(f'Succesful reply sent to {message.SenderName}.')
    else:
        logging.info(f'Sending reply of lack of attachment to {message.SenderName}.')
        reply_to_message(message, LACK_OF_ATTACHMENT_REPLY.format(message.SenderName))
        logging.info(f'Reply of lack of attachment sent to {message.SenderName}.')


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
                    send_reply(message)
                logging.info(f'All replies sent. Waiting {CHECK_INTERVAL} seconds before checking for new messages.')
                time.sleep(CHECK_INTERVAL)
            except Exception as error:
                logging.exception(f'An error {error.__class__.__name__} occurred.')
                send_message(f'Error occurred on {platform.node()}\nProcess: "Сверка зарплатной ведомости"\nError:\n{error}')
                SESSION.close()
                logging.info('SQLite session closed.')
                logging.exception(error)
                raise error
            except KeyboardInterrupt as error:
                logging.exception(f'Keyboard interrupt occurred.')
                SESSION.close()
                logging.info('SQLite session closed.')
                logging.exception(error)
                raise error
