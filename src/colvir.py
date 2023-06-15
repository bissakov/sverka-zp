import pywinauto
from pywinauto import keyboard
from time import sleep
from excel import Excel
from data_structures import Credentials, Process, ExcelInfo


class Colvir:
    def __init__(self, credentials: Credentials, process: Process, excel: ExcelInfo) -> None:
        self.credentials = credentials
        self.process_name, self.process_path = process.name, process.path
        self.pid = None
        self.desktop = pywinauto.Desktop(backend='uia')
        self.app = None
        self.excel = excel
        self.excel_name = ''
        self.excel_date = ''
        self.errored = False

    def run(self) -> None:
        self.correct_excel(self.excel)
        try:
            self._run()
        except pywinauto.findwindows.ElementNotFoundError:
            self.kill()
            self.run()
            return

        sleep(1)
        loading_import_win = self.app.window(title='Colvir Banking System', found_index=0)
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)

        sleep(2)
        if self.app.window(title='Произошла ошибка').exists():
            self.errored = True
            self.kill()
            return

        main_win = self.app.window(title='Импорт файлов', found_index=0)
        main_win.wait(wait_for='exists', timeout=60)
        main_win.set_focus()

        keyboard.send_keys('{VK_F5}')

        select_win = self.app.window(title='Выбор отчета')
        select_win.wait(wait_for='exists', timeout=60)
        select_win['Предварительный просмотр'].click()

        keyboard.send_keys('{VK_F9}')

        filter_win = self.app.window(title='Фильтр')
        filter_win.wait(wait_for='exists', timeout=60)
        filter_win['Edit2'].set_text(text='Z_160_RPT_IMP_FZDOHOD')
        filter_win['OK'].click()
        sleep(.5)

        loading_import_win = self.app.window(title='Colvir Banking System')
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)

        select_win['Экспорт в файл...'].click()

        file_win = self.app.window(title='Файл отчета ')
        file_win.wait(wait_for='exists', timeout=60)
        file_win['Edit2'].set_text(text='Z_160_RPT_IMP_FZDOHOD.xml')
        sleep(.05)
        file_win['Edit'].set_text(text=r'C:\REPORTS')
        sleep(.05)
        try:
            file_win['ComboBox'].select(8)
        except (IndexError, ValueError):
            pass

        file_win['OK'].click()

        settings_win = self.app.window(title='Параметры отчета ')
        settings_win['Edit0'].set_text(text=self.excel_date)
        settings_win['OK'].click()

        sleep(2)
        loading_import_win = self.app.window(title='Colvir Banking System')
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)
        sleep(2)
        loading_import_win = self.app.window(title='Colvir Banking System')
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)

        select_win = self.app.window(title='Выбор отчета')
        select_win.set_focus()
        select_win.wait(wait_for='active', timeout=60)
        keyboard.send_keys('{ESC}')

        self.choose_mode('TREPRT')

        select_win = self.app.window(title='Выбор отчета')
        select_win.wait(wait_for='exists', timeout=60)
        select_win['Предварительный просмотр'].click()

        keyboard.send_keys('{VK_F9}')

        filter_win = self.app.window(title='Фильтр')
        filter_win.wait(wait_for='exists', timeout=60)
        filter_win['Edit2'].set_text(text='Z_160_RPT_FZDOHOD')
        filter_win['OK'].click()
        sleep(.5)

        loading_import_win = self.app.window(title='Colvir Banking System')
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)

        select_win['Экспорт в файл...'].click()

        file_win = self.app.window(title='Файл отчета ')
        file_win.wait(wait_for='exists', timeout=60)
        file_win['Edit2'].set_text(text='Z_160_RPT_FZDOHOD.xml')
        sleep(.05)
        file_win['Edit'].set_text(text=r'C:\REPORTS')
        sleep(.05)
        try:
            file_win['ComboBox'].select(8)
        except (IndexError, ValueError):
            pass

        file_win['OK'].click()

        settings_win = self.app.window(title='Параметры отчета ')
        settings_win['Edit0'].set_text(text=self.excel_date)
        settings_win['Edit2'].set_text(text='Штатные сотрудники')
        settings_win['OK'].click()

        sleep(2)
        loading_import_win = self.app.window(title='Colvir Banking System')
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)
        sleep(2)
        loading_import_win = self.app.window(title='Colvir Banking System')
        loading_import_win.wait_not(wait_for_not='exists', timeout=360)

        sleep(1)
        self.kill()

    def _run(self) -> None:
        self.login()
        sleep(4)

        self.app = pywinauto.Application(backend='uia').connect(process=self.pid)
        self.confirm_warning()
        sleep(1)
        try:
            self.choose_mode('C_IMPFILELUSER')
        except pywinauto.findwindows.ElementNotFoundError as e:
            raise e

        main_win = self.app.window(title='Импорт файлов', found_index=0)
        main_win.wait(wait_for='exists', timeout=60)
        main_win.set_focus()

        keyboard.send_keys('{VK_F7}')
        filter_win = self.app.window(title='Поиск по наименованию')
        filter_win.wait(wait_for='exists', timeout=60)
        filter_win['Edit'].set_text('Z_160_IMP_FZ_DOHOD')
        filter_win['OK'].click()

        keyboard.send_keys('%')
        keyboard.send_keys('{ENTER}')
        keyboard.send_keys('{VK_DOWN}')
        keyboard.send_keys('{ENTER}')

        open_win = self.app.window(title='Open')
        open_win.wait(wait_for='exists', timeout=60)
        open_win['File name:Edit'].set_text(self.excel_name)

        keyboard.send_keys('{ENTER}')

    def correct_excel(self, excel: ExcelInfo) -> None:
        exc = Excel(excel)
        exc.correct()
        self.excel_name = exc.corrected_name
        self.excel_date = exc.date

    def login(self) -> None:
        pywinauto.Application(backend='uia').start(cmd_line=self.process_path)

        login_win = self.desktop.window(title='Вход в систему')
        login_win.wait(wait_for='exists', timeout=60)

        self.pid = login_win.wrapper_object().process_id()

        login_win.Edit2.set_text(text=self.credentials.usr)
        login_win.Edit0.set_text(text=self.credentials.psw)

        login_win.OK.click()

    def confirm_warning(self) -> None:
        win = self.app.top_window()
        button = win['Button']
        if button.exists(timeout=2):
            button.click()

    def choose_mode(self, mode: str) -> None:
        mode_win = self.app.window(title='Выбор режима')
        mode_win['Edit'].set_text(text=mode)
        keyboard.send_keys('{ENTER}')

    def kill(self) -> None:
        self.app.kill()
