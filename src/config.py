import os
import dotenv
import requests
from requests.adapters import HTTPAdapter

dotenv.load_dotenv()

PROJECT_FOLDER = r'C:\Users\robot.ad\Desktop\sverka-zp'
EXCEL_FOLDER = r'C:\Users\robot.ad\Desktop\sverka-zp\excel_reports'
CHECK_INTERVAL: int = 60
# RECIPIENTS: list[str] = ['robot.ad']
SUBJECT: str = 'Сверка зарплатной ведомости'

# REQUIRED_FILE_FORMAT: str = '.xlsx'

reply_dummy_message: str = 'Добрый день, {}\n\n' \
                           '{}\n\nПросьба не отвечать на это письмо.\n\n' \
                           'Сообщение сгенерировано автоматически.'
REPLY_MESSAGE: str = reply_dummy_message.format('{}', 'Ответ от робота.')
LACK_OF_ATTACHMENT_REPLY: str = reply_dummy_message.format('{}', 'Отсутствует вложенный файл.\n'
                                                                 'Пожалуйста приложите файл и '
                                                                 'отправьте новое отдельное письмо.')
ATTACHMENTS_MORE_THAN_ONE_REPLY: str = reply_dummy_message.format('{}', 'Вложено больше одного файла.\n'
                                                                        'Пожалуйста приложите только один файл '
                                                                        'и отправьте новое отдельное письмо.')
WRONG_ATTACHMENT_FORMAT_REPLY: str = reply_dummy_message.format('{}', 'Неверный формат вложенного файла.\n'
                                                                      'Пожалуйста приложите файл в формате .xlsx '
                                                                      'и отправьте новое отдельное письмо.')

BOT_TOKEN = os.getenv('BOT_TOKEN')
CHAT_ID = os.getenv('CHAT_ID')
SESSION = requests.Session()
SESSION.mount('http://', HTTPAdapter(max_retries=5))


# excel_path = r'\\dbu00234\c$\Temp\Сверка'
# excel_name = [f for f in listdir(excel_path) if isfile(join(excel_path, f))][0]
#
# excel_path = excel_path
# excel_name = excel_name
# email_list = 'zhekenova.a; abdullayeva.b@otbasybank.kz;' \
#              'baktibay.d@otbasybank.kz; abdieva.g@otbasybank.kz; robot.ad@hcsbkkz.loc'
