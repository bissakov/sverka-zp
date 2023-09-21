import logging
import os
import shutil
from time import sleep
from datetime import date
from os.path import join

import win32com.client as win32
from sqlalchemy import Column, ColumnElement, create_engine, String
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

import colvir
import excel
import logger
from config import ATTACHMENTS_MORE_THAN_ONE_REPLY, CHECK_INTERVAL, EXCEL_FOLDER, LACK_OF_ATTACHMENT_REPLY, \
    REPLY_MESSAGE, SUBJECT
from utils import dispatch

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
Base.metadata.create_all(engine)


def clean_database():
    with sessionmaker(bind=engine)() as session:
        session.query(Reply).delete()
        session.commit()


def save_reply(message_id: str) -> None:
    with sessionmaker(bind=engine)() as session:
        replied_email = Reply(message_id=message_id)
        session.add(replied_email)
        session.commit()


def get_replied_messages() -> list[ColumnElement]:
    with sessionmaker(bind=engine)() as session:
        replied_emails = session.query(Reply.message_id).all()
    return [email[0] for email in replied_emails]


def get_messages(outlook: win32.CDispatch) -> list[win32.CDispatch]:
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
            message.EntryID not in get_replied_messages()]


def reply_to_message(message: win32.CDispatch, reply_message: str, attachment: str = None) -> None:
    reply = message.Reply()
    reply.Body = reply_message
    if attachment:
        attachment = f'{attachment}.zip' if 'zip' not in attachment else attachment
        reply.Attachments.Add(attachment)
        logging.info(f'Attached file "{attachment}".')
    reply.Send()
    save_reply(message.EntryID)
    logging.info(f'Saved message id "{message.EntryID}" to replied emails file.')


def validate_message(message: win32.CDispatch) -> tuple[bool, str, str]:
    attachment_count = message.Attachments.Count
    sender_name = message.SenderName
    if attachment_count == 0:
        return False, 'LACK_OF_ATTACHMENT_REPLY', LACK_OF_ATTACHMENT_REPLY.format(sender_name)
    elif attachment_count > 1:
        return False, 'ATTACHMENTS_MORE_THAN_ONE_REPLY', ATTACHMENTS_MORE_THAN_ONE_REPLY.format(sender_name)
    else:
        return True, 'REPLY_MESSAGE', REPLY_MESSAGE.format(message.SenderName)


def save_attachment(message: win32.CDispatch) -> str:
    attachment = message.Attachments.Item(1)
    excel_name_to_correct = attachment.FileName
    excel_to_correct = join(EXCEL_FOLDER, excel_name_to_correct)
    attachment.SaveAsFile(excel_to_correct)
    if not excel_to_correct.endswith('xlsx'):
        with dispatch('Excel.Application') as excel_app:
            excel_app.Workbooks.Open(excel_to_correct)
            extension = excel_to_correct.split('.')[-1]
            excel_app.ActiveWorkbook.SaveAs(excel_to_correct.replace(f'.{extension}', '.xlsx'), FileFormat=51)
            excel_app.ActiveWorkbook.Close(True)
            os.unlink(excel_to_correct)
            excel_name_to_correct = excel_name_to_correct.replace(f'.{extension}', '.xlsx')
    logging.info(f'Saved attachment "{excel_name_to_correct}" to "{EXCEL_FOLDER}" folder.')
    return excel_name_to_correct


def make_archive(exports_folder: str) -> str:
    zip_file = join(EXCEL_FOLDER, 'протокол_ошибок')
    with dispatch('Excel.Application') as excel_app:
        for file in os.listdir(exports_folder):
            full_file_path = join(exports_folder, file)
            excel_app.Workbooks.Open(full_file_path)
            excel_app.ActiveWorkbook.SaveAs(full_file_path.replace('.xls', ''), FileFormat=50)
            excel_app.ActiveWorkbook.Close(True)
            os.unlink(full_file_path)
    shutil.make_archive(zip_file, 'zip', exports_folder)
    return f'{zip_file}.zip'


def run(check_interval: int = CHECK_INTERVAL) -> None:
    with dispatch('Outlook.Application') as outlook_namespace:
        logging.info('Outlook application started.')
        while True:
            logging.info('Checking inbox for new messages.')
            messages = get_messages(outlook_namespace)

            if not messages:
                logging.info(f'No new messages. Waiting {check_interval} seconds before checking inbox.')
                sleep(check_interval)
                continue

            logging.info(f'Found {len(messages)} new messages.')

            for message in messages:
                is_valid, reply_type, reply_message = validate_message(message)
                if not is_valid:
                    logging.info(f'Sending {reply_type} reply to {message.SenderName}.')
                    reply_to_message(message, reply_message)
                    logging.info(f'Reply {reply_type} sent to {message.SenderName}.')
                else:
                    logging.info('Starting the process "Сверка зарплатной ведомости".')

                    excel_name_to_correct = save_attachment(message)
                    corrected_excel_name, excel_date = excel.correct(excel_name=excel_name_to_correct)
                    exports_folder = join(EXCEL_FOLDER, rf'exports\{excel_date}')
                    os.makedirs(exports_folder, exist_ok=True)
                    colvir.run(corrected_excel_name, exports_folder)
                    zip_file = make_archive(exports_folder)

                    logging.info(f'Sending {reply_type} reply to {message.SenderName}.')
                    reply_to_message(message, reply_message, attachment=zip_file)
                    logging.info(f'Reply {reply_type} sent to {message.SenderName}.')

            logging.info(f'All replies sent. Waiting {check_interval} seconds before checking inbox.')
            sleep(check_interval)
