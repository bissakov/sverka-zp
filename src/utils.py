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


def choose_mode(mode: str, app: Application | None = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=r'C:CBS_R_NEWCBS_RCOLVIR.EXE')
    mode_win = app.window(title='Выбор режима')
    mode_win['Edit2'].set_text(text=mode)
    mode_win['Edit2'].send_keystrokes('~')
    # press(mode_win['Edit2'], '~')


def is_errored(app: Application) -> bool:
    for win in app.windows():
        text = win.window_text().strip()
        if text and 'Произошла ошибка' in text:
            return True
    return False


def is_correct_file(excel_full_file_path: str, excel: win32.Dispatch) -> bool:
    extension = excel_full_file_path.split('.')[-1]
    excel_full_file_path_no_ext = '.'.join(excel_full_file_path.split('.')[0:-1])
    excel_copy_path = f'{excel_full_file_path_no_ext}_copy.{extension}'
    shutil.copyfile(src=excel_full_file_path, dst=excel_copy_path)
    xlsx_file_path = f'{excel_full_file_path_no_ext}.xlsx'

    if not exists(path=xlsx_file_path):
        wb = excel.Workbooks.Open(excel_copy_path)
        wb.SaveAs(xlsx_file_path, FileFormat=51)
        wb.Close()

    workbook: Workbook = openpyxl.load_workbook(xlsx_file_path, data_only=True)
    sheet: Worksheet = workbook.active
    unlink(xlsx_file_path)
    unlink(excel_copy_path)

    return next((True for row in sheet.iter_rows(max_row=50) for cell in row if cell.alignment.horizontal), False)


def type_keys(window: WindowSpecification, keystrokes: str, step_delay: float = .1) -> None:
    set_focus(window)
    for command in list(filter(None, re.split(r'({.+?})', keystrokes))):
        try:
            window.type_keys(command)
        except ElementNotEnabled:
            sleep(1)
            window.type_keys(command)
        sleep(step_delay)


def press(win: WindowSpecification, key: str, pause: float = 0) -> None:
    set_focus(win)
    win.type_keys(key, pause=pause)


def set_focus(win: WindowSpecification) -> None:
    while not win.is_active():
        try:
            win.set_focus()
            break
        except Exception as error:
            _ = error
            sleep(1)
            continue


def click_input(win: WindowSpecification, coords: tuple[int, int] = None, delay: float = 0) -> None:
    sleep(delay)
    set_focus(win)
    sleep(delay)
    if coords:
        win.click_input(button='left', coords=coords, absolute=True)
    else:
        win.click_input()


def double_click_input(win: WindowSpecification, delay: float = 0) -> None:
    sleep(delay)
    set_focus(win)
    sleep(delay)
    win.double_click_input()
