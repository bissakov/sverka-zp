import os
import psutil
from time import sleep
from datetime import datetime as dt
import win32com.client as win32
import shutil
from archive import Archive
from mail import Email
from colvir import Colvir
from data_structures import Credentials, Process, ExcelInfo, ArchiveInfo, EmailInfo


class Robot:
    def __init__(self, credentials: Credentials, process: Process, excel: ExcelInfo,
                 archive_info: ArchiveInfo, email_info: EmailInfo) -> None:
        self.colvir = Colvir(
            credentials=credentials,
            process=process,
            excel=excel
        )
        self.archive_info = archive_info
        self.process_name, self.dir = process.name, archive_info.zip_dir
        self.archive = Archive(info=archive_info)
        self.today = dt.now().strftime('%d.%m.%Y')
        self.email_info = email_info
        self.attachment = archive_info.zip_name
        self.excel = excel

    def run(self) -> None:
        self.kill_all()
        self.colvir.run()
        if self.colvir.errored:
            email = Email(
                email_info=self.email_info,
                subject=f'Возникла ошибка при импорте файла. '
                        f'Просьба перепроверить оригинальный файл {self.today}',
                body=f'Возникла ошибка при импорте файла. '
                     f'Пожалуйста перепроверьте оригинальный файл.\n\nСообщение сгенерировано автоматически. '
                     f'Просьба не отвечать.',
            )
            email.run()
            self.return_file()
            return
        self.open_folder()
        self.convert_to_xlsb()
        self.archive.run()
        email = Email(
            email_info=self.email_info,
            subject=f'Протокол ошибок по оплате труда {self.today}',
            body=f'Вложенные протоколы по работе робота.\n\n'
                 f'Сообщение сгенерировано автоматически. Просьба не отвечать.',
            attachment=self.attachment
        )
        email.run()
        self.return_file()

    def convert_to_xlsb(self):
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        for report_file in os.listdir(self.archive_info.zip_dir):
            file_name = os.path.join(self.archive_info.zip_dir, report_file)
            if not os.path.isfile(file_name):
                continue
            wb = excel.Workbooks.OpenXML(file_name)
            wb.SaveAs(file_name.replace('.xml', '.xlsb'), 50)
            wb.Close()
            os.unlink(file_name)
        excel.Quit()

    def open_folder(self) -> None:
        path = os.path.realpath(self.dir)
        os.startfile(path)

    def kill_all(self) -> None:
        for proc in psutil.process_iter():
            if any(process_name in proc.name() for process_name in [self.process_name, 'EXCEL']):
                try:
                    p = psutil.Process(proc.pid)
                    p.terminate()
                except psutil.AccessDenied:
                    continue
        sleep(.5)

    def return_file(self):
        source = os.path.join(self.excel.path, self.excel.name)
        destination = os.path.join(self.excel.path.replace(r'\Сверка', ''), self.excel.name)
        shutil.move(source, destination)
