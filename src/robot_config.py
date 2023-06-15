import os
from os import listdir
from os.path import isfile, join
from data_structures import Credentials, Process, ExcelInfo, ArchiveInfo, EmailInfo
import dotenv

dotenv.load_dotenv()

excel_path = r'\\dbu00234\c$\Temp\Сверка'
excel_name = [f for f in listdir(excel_path) if isfile(join(excel_path, f))][0]

excel_path = excel_path
excel_name = excel_name
email_list = 'zhekenova.a; abdullayeva.b@otbasybank.kz;' \
             'baktibay.d@otbasybank.kz; abdieva.g@otbasybank.kz; robot.ad@hcsbkkz.loc'

CREDENTIALS = Credentials(usr=os.getenv('COLVIR_USER'), psw=os.getenv('COLVIR_PASSWORD'))
PROCESS = Process(
    name='COLVIR',
    path=r'C:\CBS_R_NEW\CBS_R\COLVIR.EXE'
)
EXCEL = ExcelInfo(path=excel_path, name=excel_name)
EMAIL_INFO = EmailInfo(email_list=email_list)
