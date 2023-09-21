import logging
import os
import re
from os import unlink
from os.path import exists, getsize, join
from time import sleep

import dotenv
import pywinauto
import win32com.client as win32
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError

from config import EXCEL_FOLDER
from data_structures import Credentials
from utils import choose_mode, click_input, dispatch, get_window, is_correct_file, kill_all_processes, type_keys

dotenv.load_dotenv()
CREDENTIALS = Credentials(user=os.getenv('COLVIR_USER'), password=os.getenv('COLVIR_PASSWORD'))


def get_app(title: str, backend: str = 'win32') -> Application:
    app = None
    while not app:
        try:
            app = Application(backend=backend).connect(title=title)
        except ElementNotFoundError:
            sleep(.1)
            continue
    return app


def login(app: Application | None = None) -> None:
    if not app:
        app = get_app(title='Вход в систему')
    login_win = app.window(title='Вход в систему')
    login_win['Edit2'].set_text(text=CREDENTIALS.user)
    login_win['Edit'].set_text(text=CREDENTIALS.password)
    login_win['OK'].click()


def confirm(app: Application | None = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=r'C:\CBS_R\COLVIR.EXE')
    dialog = app.window(title='Colvir Banking System', found_index=0)
    timeout = 0
    while not dialog.window(best_match='OK').exists():
        if timeout >= 5.0:
            raise pywinauto.findwindows.ElementNotFoundError
        timeout += .1
        sleep(.1)
    dialog.send_keystrokes('~')


def check_interactivity(app: Application | None = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=r'C:\CBS_R\COLVIR.EXE')
    choose_mode(app=app, mode='EXTRCT')
    sleep(1)
    if (filter_win := app.window(title='Фильтр')).exists():
        filter_win.close()
    else:
        raise pywinauto.findwindows.ElementNotFoundError


def open_colvir() -> Application:
    retry_count: int = 0
    app = None
    while retry_count < 5:
        try:
            app = Application().start(cmd_line=r'C:\CBS_R\COLVIR.EXE')
            login()
            confirm()
            check_interactivity()
            break
        except pywinauto.findwindows.ElementNotFoundError:
            retry_count += 1
            if app:
                app.kill()
            continue
    if retry_count == 5:
        raise Exception('max_retries exceeded')
    return app


def import_excel(app: Application, excel_name: str) -> None:
    choose_mode(app=app, mode='C_IMPFILELUSER')
    main_win = get_window(title='Импорт файлов', app=app)

    type_keys(main_win, '{VK_F7}', step_delay=.5)
    filter_win = get_window(title='Поиск по наименованию', app=app)
    type_keys(filter_win['Edit2'], 'Z_160_IMP_FZ_DOHOD', step_delay=.5)
    click_input(filter_win['OK'])

    type_keys(main_win,
              '{VK_SHIFT down}{VK_MENU}н{VK_SHIFT up}{DOWN}~',
              step_delay=.5)

    open_win = get_window(title='Open', app=app)
    type_keys(open_win['File name:Edit'], join(EXCEL_FOLDER, excel_name), step_delay=.5)
    while open_win['Open'].exists():
        click_input(open_win['Open'])
        sleep(1)

    opened_windows_count = len(app.windows())
    current_windows_count = opened_windows_count
    logging.info(f'Current windows count after starting export: {current_windows_count}')

    while current_windows_count == opened_windows_count:
        current_windows_count = len(app.windows())
        sleep(5)
    logging.info(f'Current windows count: {current_windows_count}')


def is_file_exported(full_file_name: str, excel: win32.CDispatch) -> bool:
    if not exists(path=full_file_name):
        logging.info(f'File {full_file_name} does not yet exist')
        return False
    if getsize(filename=full_file_name) == 0:
        logging.info(f'File {full_file_name} is empty')
        return False
    try:
        os.rename(src=full_file_name, dst=full_file_name)
        logging.info(f'File {full_file_name} is not locked by Colvir')
    except OSError:
        logging.info(f'File {full_file_name} is still being written by Colvir')
        return False
    if not is_correct_file(excel_full_file_path=full_file_name, excel=excel):
        logging.info(f'File {full_file_name} has no horizional alignment the first 50 rows.')
        return False
    logging.info(f'File {full_file_name} exists and ready')
    return True


def export(app: Application, exports_folder: str,  report_type: str) -> None:
    report_win = get_window(app=app, title='Выбор отчета')
    sleep(1)
    type_keys(report_win, '{F9}', step_delay=.5)

    sleep(1)
    filter_win = get_window(app=app, title='Фильтр')
    type_keys(filter_win['Edit4'], report_type, step_delay=.5)
    sleep(1)
    click_input(filter_win['OK'])
    sleep(1)

    click_input(report_win['Предварительный просмотр'])
    export_win = report_win['Экспорт в файл...']
    export_win.wait(wait_for='enabled', timeout=30)

    click_input(export_win)

    file_name = f'{report_type}.xls'
    full_file_name = join(exports_folder, file_name)

    if exists(full_file_name):
        unlink(full_file_name)

    file_win = get_window(app=app, title='Файл отчета ')
    type_keys(file_win['Edit2'], exports_folder, step_delay=.5)
    sleep(1)
    type_keys(file_win['Edit4'], file_name, step_delay=.5)
    try:
        file_win['ComboBox'].select(11)
        sleep(1)
    except (IndexError, ValueError):
        pass
    click_input(file_win['OK'])

    params_win = get_window(app=app, title='Параметры отчета ')

    excel_date = exports_folder.split('\\')[-1]
    assert re.match(r'[A-Z]{3}_\d{4}', excel_date), f'Excel name {excel_date} does not match pattern [A-Z]{{3}}_\\d{{4}}'
    type_keys(params_win['Edit2'], excel_date, step_delay=.5)
    if report_type == 'Z_160_RPT_FZDOHOD':
        type_keys(params_win['Edit4'], 'Ш', step_delay=.5)
    click_input(params_win['OK'])

    sleep(30)
    with dispatch(application='Excel.Application') as excel:
        while not is_file_exported(full_file_name=full_file_name, excel=excel):
            sleep(5)
    report_win.close()


def export_z_160_rpt_imp_fzdohod(app: Application, exports_folder: str) -> None:
    main_win = get_window(title='Импорт файлов', app=app)
    type_keys(main_win, '{F5}', step_delay=.5)
    export(app=app, exports_folder=exports_folder, report_type='Z_160_RPT_IMP_FZDOHOD')
    main_win.close()


def export_z_160_rpt_fzdohod(app: Application, exports_folder: str):
    choose_mode(mode='TREPRT', app=app)
    export(app=app, exports_folder=exports_folder, report_type='Z_160_RPT_FZDOHOD')


def run(excel_name: str, exports_folder: str) -> None:
    logging.info('Colvir process started.')

    kill_all_processes(proc_name='COLVIR')
    logging.info('All previous Colvir process killed.')

    app = open_colvir()
    logging.info('Colvir opened and logged in.')

    import_excel(app=app, excel_name=excel_name)
    logging.info('Excel imported.')

    logging.info('Z_160_RPT_IMP_FZDOHOD export started.')
    export_z_160_rpt_imp_fzdohod(app=app, exports_folder=exports_folder)
    logging.info('Z_160_RPT_IMP_FZDOHOD exported.')

    logging.info('Z_160_RPT_FZDOHOD export started.')
    export_z_160_rpt_fzdohod(app=app, exports_folder=exports_folder)
    logging.info('Z_160_RPT_FZDOHOD exported.')

    kill_all_processes(proc_name='COLVIR')
    logging.info('All Colvir processes killed.')

    logging.info('Colvir process finished.')
