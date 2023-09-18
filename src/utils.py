import re
import shutil
from contextlib import contextmanager
from os import unlink
from os.path import exists, join
from time import sleep
from typing import List, Optional

import openpyxl
import psutil
import win32com.client as win32
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from psutil import Process
from pywinauto import Application, WindowSpecification
from pywinauto.base_wrapper import ElementNotEnabled


@contextmanager
def dispatch(application: str) -> None:
    if 'Outlook' in application and get_current_process_pid(proc_name='OUTLOOK'):
        kill_all_processes(proc_name='OUTLOOK')
    if 'Excel' in application and get_current_process_pid(proc_name='EXCEL'):
        kill_all_processes(proc_name='EXCEL')
    app = win32.Dispatch(application)
    namespace = app.GetNamespace('MAPI') if 'Outlook' in application else None
    if 'Excel' in application:
        app.DisplayAlerts = False
    try:
        yield namespace if namespace else app
    finally:
        app.Quit()


def kill_process(pid: Optional[int]) -> None:
    if pid is None:
        raise ValueError('pid is None')
    proc = Process(pid)
    proc.terminate()


def kill_all_processes(proc_name: str) -> None:
    processes_to_kill: List[Process] = [Process(proc.pid) for proc in psutil.process_iter() if
                                        proc_name in proc.name()]
    for process in processes_to_kill:
        try:
            process.terminate()
        except psutil.AccessDenied:
            continue


def get_current_process_pid(proc_name: str) -> int or None:
    return next((p.pid for p in psutil.process_iter() if proc_name in p.name()), None)


def get_window(title: str, app: Application, wait_for: str = 'exists', timeout: int = 20,
               regex: bool = False, found_index: int = 0) -> WindowSpecification:
    window = app.window(title=title, found_index=found_index) \
        if not regex else app.window(title_re=title, found_index=found_index)
    window.wait(wait_for=wait_for, timeout=timeout)
    sleep(.5)
    return window


def type_keys(window, keystrokes: str, step_delay: float = .1) -> None:
    for command in list(filter(None, re.split(r'({.+?})', keystrokes))):
        try:
            window.type_keys(command)
        except ElementNotEnabled:
            sleep(1)
            window.type_keys(command)
        sleep(step_delay)


def choose_mode(app: Application, mode: str) -> None:
    mode_win = app.window(title='Выбор режима')
    mode_win.wait(wait_for='exists', timeout=60)
    mode_win['Edit2'].wrapper_object().set_text(text=mode)
    mode_win['Edit2'].wrapper_object().send_keystrokes('~')


def is_errored(app: Application) -> bool:
    for win in app.windows():
        text = win.window_text().strip()
        if text and 'Произошла ошибка' in text:
            return True
    return False


def is_correct_file(root: str, xls_file_path: str, excel: win32.Dispatch) -> bool:
    xls_file_path = join(root, xls_file_path)
    shutil.copyfile(src=xls_file_path, dst=f'{xls_file_path}_copy.xls')
    xls_file_path = f'{xls_file_path}_copy.xls'
    xlsx_file_path = xls_file_path + 'x'

    if not exists(path=xlsx_file_path):
        wb = excel.Workbooks.Open(xls_file_path)
        wb.SaveAs(xlsx_file_path, FileFormat=51)
        wb.Close()

    workbook: Workbook = openpyxl.load_workbook(xlsx_file_path, data_only=True)
    sheet: Worksheet = workbook.active
    unlink(xlsx_file_path)
    unlink(xls_file_path)

    return next((True for row in sheet.iter_rows(max_row=50) for cell in row if cell.alignment.horizontal), False)
