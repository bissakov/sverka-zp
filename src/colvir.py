import logging
import os
from os import unlink
from os.path import exists, getsize, join
from time import sleep

import dotenv
import pywinauto
import win32com.client as win32
from pywinauto import Application, Desktop
from pywinauto.application import ProcessNotFoundError
from pywinauto.controls.hwndwrapper import DialogWrapper
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
from pywinauto.timings import TimeoutError as TimingsTimeoutError

from config import EXCEL_FOLDER
from data_structures import Credentials, Process
from utils import choose_mode, dispatch, get_current_process_pid, get_window, is_correct_file, is_errored, \
    kill_all_processes, kill_process, type_keys

dotenv.load_dotenv()
CREDENTIALS = Credentials(usr=os.getenv('COLVIR_USER'), psw=os.getenv('COLVIR_PASSWORD'))
PROCESS = Process(name='COLVIR', path=r'C:\CBS_R_NEW\CBS_R\COLVIR.EXE')


def login() -> None:
    desktop: Desktop = Desktop(backend='win32')
    try:
        login_win = desktop.window(title='Вход в систему')
        login_win.wait(wait_for='exists', timeout=20)
        login_win['Edit2'].wrapper_object().set_text(text=CREDENTIALS.usr)
        login_win['Edit'].wrapper_object().set_text(text=CREDENTIALS.psw)
        login_win['OK'].wrapper_object().click()
    except ElementAmbiguousError:
        windows: list[DialogWrapper] = Desktop(backend='win32').windows()
        for win in windows:
            if 'Вход в систему' not in win.window_text():
                continue
            kill_process(pid=win.process_id())
        raise ElementNotFoundError


def confirm_warning(app: Application) -> None:
    found = False
    for window in app.windows():
        if found:
            break
        if window.window_text() != 'Colvir Banking System':
            continue
        win = app.window(handle=window.handle)
        for child in win.descendants():
            if child.window_text() == 'OK':
                found = True
                win.close()
                if win.exists():
                    win.close()
                break


def open_colvir(retry_count: int = 0) -> Application | None:
    if retry_count == 3:
        raise RuntimeError('Не удалось запустить Colvir')

    try:
        try:
            Application(backend='win32').start(cmd_line=PROCESS.path)
        except pywinauto.application.AppStartError:
            # TODO change in production
            Application(backend='win32').start(cmd_line=r'C:\CBS_R\COLVIR.EXE')
        login()
        sleep(4)
    except (ElementNotFoundError, TimingsTimeoutError):
        retry_count += 1
        kill_all_processes(proc_name=PROCESS.name)
        app = open_colvir(retry_count=retry_count)
        return app
    try:
        pid: int = get_current_process_pid(proc_name=PROCESS.name)
        app: Application = Application(backend='win32').connect(process=pid)
        try:
            if app.Dialog.window_text() == 'Произошла ошибка':
                retry_count += 1
                kill_all_processes(proc_name=PROCESS.name)
                app = open_colvir(retry_count=retry_count)
                return app
        except MatchError:
            pass
    except ProcessNotFoundError:
        sleep(1)
        pid = get_current_process_pid(proc_name='COLVIR')
        app: Application = Application(backend='win32').connect(process=pid)
    try:
        confirm_warning(app=app)
        sleep(2)
        if is_errored(app=app):
            raise ElementNotFoundError
    except (ElementNotFoundError, MatchError):
        retry_count += 1
        kill_all_processes(proc_name=PROCESS.name)
        app = open_colvir(retry_count=retry_count)
    return app


def import_excel(app: Application, excel_name: str) -> None:
    choose_mode(app=app, mode='C_IMPFILELUSER')
    main_win = get_window(title='Импорт файлов', app=app)

    main_win.send_keystrokes('{VK_F7}')
    filter_win = get_window(title='Поиск по наименованию', app=app)
    filter_win['Edit2'].set_text('Z_160_IMP_FZ_DOHOD')
    filter_win['OK'].click()

    type_keys(main_win,
              '{VK_SHIFT down}{VK_MENU}н{VK_SHIFT up}{DOWN}~',
              step_delay=.5)

    open_win = get_window(title='Open', app=app)
    open_win['File name:Edit'].set_text(join(EXCEL_FOLDER, excel_name))
    while open_win['Open'].exists():
        open_win['Open'].click()
        sleep(1)

    opened_windows_count = len(app.windows())
    current_windows_count = opened_windows_count
    logging.info(f'Current windows count after starting export: {current_windows_count}')

    while current_windows_count == opened_windows_count:
        current_windows_count = len(app.windows())
        logging.info(f'Current windows count: {current_windows_count}')
        sleep(30)


def is_file_exported(file_name: str, excel: win32.CDispatch) -> bool:
    full_file_name = join(EXCEL_FOLDER, 'exports', file_name)
    if not exists(path=full_file_name):
        logging.info(f'File {full_file_name} does not yet exist')
        return False
    if getsize(filename=full_file_name) == 0:
        logging.info(f'File {full_file_name} is empty')
        return False
    try:
        os.rename(src=full_file_name, dst=full_file_name)
        logging.info(f'File {full_file_name} is not locked by Excel')
    except OSError:
        logging.info(f'File {full_file_name} is still being written by Excel')
        return False
    if not is_correct_file(root=join(EXCEL_FOLDER, 'exports'), xls_file_path=file_name, excel=excel):
        logging.info(f'File {full_file_name} has no horizional alignment the first 50 rows.')
        return False
    logging.info(f'File {full_file_name} exists and ready')
    return True


def export(app: Application, excel_date: str,  report_type: str) -> None:
    report_win = get_window(app=app, title='Выбор отчета')
    sleep(1)
    report_win.send_keystrokes('{F9}')

    sleep(1)
    filter_win = get_window(app=app, title='Фильтр')
    filter_win['Edit4'].set_text(text=report_type)
    sleep(1)
    filter_win['OK'].wrapper_object().click()
    sleep(1)

    report_win['Предварительный просмотр'].wrapper_object().click()
    export_win = report_win['Экспорт в файл...']
    export_win.wait(wait_for='enabled', timeout=30)

    export_win.wrapper_object().click()

    file_name = f'{report_type}.xml'
    full_file_name = join(EXCEL_FOLDER, 'exports', file_name)

    if exists(full_file_name):
        unlink(full_file_name)

    file_win = get_window(app=app, title='Файл отчета ')
    file_win['Edit2'].set_text(text=join(EXCEL_FOLDER, 'exports'))
    sleep(1)
    file_win['Edit4'].set_text(text=file_name)
    try:
        # file_win['ComboBox'].select(11 if report_type == 'Z_160_RPT_IMP_FZDOHOD' else 7)
        file_win['ComboBox'].select(7)
        sleep(1)
    except (IndexError, ValueError):
        pass
    file_win['OK'].click()

    params_win = get_window(app=app, title='Параметры отчета ')

    params_win['Edit2'].set_text(text=excel_date)
    if report_type == 'Z_160_RPT_FZDOHOD':
        params_win['Edit4'].set_text(text='Штатные сотрудники')
    params_win['OK'].click()

    sleep(30)
    with dispatch(application='Excel.Application') as excel:
        while not is_file_exported(file_name=file_name, excel=excel):
            sleep(30)
    report_win.close()


def export_z_160_rpt_imp_fzdohod(app: Application, excel_date: str) -> None:
    main_win = get_window(title='Импорт файлов', app=app)
    main_win.send_keystrokes('{F5}')
    export(app=app, excel_date=excel_date, report_type='Z_160_RPT_IMP_FZDOHOD')
    main_win.close()


def export_z_160_rpt_fzdohod(app: Application, excel_date: str):
    choose_mode(mode='TREPRT', app=app)
    export(app=app, excel_date=excel_date, report_type='Z_160_RPT_FZDOHOD')


def run(excel_name: str, excel_date: str) -> None:
    logging.info('Colvir process started.')

    kill_all_processes(proc_name=PROCESS.name)
    logging.info('All previous Colvir process killed.')

    app = open_colvir()
    logging.info('Colvir opened and logged in.')

    import_excel(app=app, excel_name=excel_name)
    logging.info('Excel imported.')

    logging.info('Z_160_RPT_IMP_FZDOHOD export started.')
    export_z_160_rpt_imp_fzdohod(app=app, excel_date=excel_date)
    logging.info('Z_160_RPT_IMP_FZDOHOD exported.')

    logging.info('Z_160_RPT_FZDOHOD export started.')
    export_z_160_rpt_fzdohod(app=app, excel_date=excel_date)
    logging.info('Z_160_RPT_FZDOHOD exported.')

    kill_all_processes(proc_name=PROCESS.name)
    logging.info('All Colvir processes killed.')

    logging.info('Colvir process finished.')
