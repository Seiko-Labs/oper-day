import os
import re
import shutil
from itertools import islice
from time import sleep
from typing import Dict, List

import openpyxl
import psutil
import pywinauto
import win32com.client as win32
from psutil import Process
from pywinauto.base_wrapper import ElementNotEnabled
from pywinauto.timings import Timings

from excel_converter import ExcelConverter


class BackendManager:
    def __init__(self, app: pywinauto.Application, backend_name: str) -> None:
        if backend_name not in ['uia', 'win32']:
            raise ValueError('backend_name must be either \'uia\' or \'win32\'')
        self.app, self.backend_name = app, backend_name

    def __enter__(self) -> None:
        self.app.backend.name = self.backend_name

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.app.backend.name = 'win32' if self.backend_name == 'uia' else 'uia'


class TimingManager:
    def __init__(self, timing: str) -> None:
        if timing not in ['fast', 'slow', 'defaults']:
            raise ValueError('timing must be either \'fast\', \'slow\' or \'defaults\'')
        self.timing = timing

    def __enter__(self) -> None:
        if self.timing == 'slow':
            Timings.slow()
        elif self.timing == 'fast':
            Timings.fast()
        else:
            Timings.defaults()

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        Timings.fast()


class Utils:
    def __init__(self) -> None:
        self.excel_converter: ExcelConverter = ExcelConverter()

    def convert(self, src_file: str, dst_file: str, file_type: str) -> None:
        self.excel_converter.convert(src_file=src_file, dst_file=dst_file, file_type=file_type)

    @staticmethod
    def kill_process(pid: int) -> None:
        proc = Process(pid)
        proc.terminate()

    @staticmethod
    def kill_all_processes(proc_name: str, restricted_pids: List[int] or None = None) -> None:
        processes_to_kill: List[Process] = [Process(proc.pid) for proc in psutil.process_iter() if
                                            proc_name in proc.name()]
        for process in processes_to_kill:
            try:
                process.terminate()
            except psutil.AccessDenied:
                if restricted_pids:
                    restricted_pids.append(process.pid)
                continue

    @staticmethod
    def get_current_process_pid(proc_name: str) -> int or None:
        return next((p.pid for p in psutil.process_iter() if proc_name in p.name()), None)

    @staticmethod
    def is_active(app) -> bool:
        try:
            return app.active()
        except RuntimeError:
            return False

    @staticmethod
    def text_to_dicts(file_path: str) -> List[Dict[str, str]]:
        pattern = re.compile(r'(Начало|Конец) записи \d+\.\d+\.\d+ \d+:\d+:\d+')
        encoding = 'utf-8' if file_path.endswith('.txt') else 'utf-16'
        with open(file=file_path, mode='r', encoding=encoding) as file:
            rows = [[el.replace('\n', '') for el in line.split('\t')] for line in file if not pattern.search(line)]
        header = [col.strip() for col in rows[0]]
        data_rows = islice(rows, 1, None)
        return [{col: val.strip() for col, val in zip(header, row)} for row in data_rows]

    @staticmethod
    def is_reg_procedure_ready(file_path: str, reg_num: str, delay: int = 5) -> bool:
        data = Utils.text_to_dicts(file_path=file_path)
        if not data:
            sleep(delay)
            return False

        user_name = 'Создатель базы данных'  # temporary

        reg = next((row for row in data if row['Исполнитель'] == user_name and row['Операция'] == f'Регламентная процедура номер {reg_num}'), None)

        if reg:
            sleep(delay)
            return False
        return True

    @staticmethod
    def is_correct_file(root: str, xls_file_name: str) -> bool:
        xls_file_path = os.path.join(root, xls_file_name)
        if not os.path.exists(xls_file_path):
            return False
        shutil.copyfile(src=xls_file_path, dst=f'{xls_file_path}_copy.xls')
        xls_file_path = f'{xls_file_path}_copy.xls'
        xlsx_file_path = xls_file_path + 'x'

        excel = win32.Dispatch('Excel.Application')
        if not os.path.exists(path=xlsx_file_path):
            wb = excel.Workbooks.Open(xls_file_path)
            wb.SaveAs(xlsx_file_path, FileFormat=51)
            wb.Close()
        excel.Quit()

        workbook = openpyxl.load_workbook(xlsx_file_path, data_only=True)
        sheet = workbook.active
        os.unlink(xlsx_file_path)
        os.unlink(xls_file_path)

        return next((True for row in sheet.iter_rows(max_row=50) for cell in row if cell.has_style), False)

    @staticmethod
    def type_keys(_window, keystrokes: str, step_delay: float = .1) -> None:
        for command in list(filter(None, re.split(r'({.+?})', keystrokes))):
            try:
                _window.type_keys(command)
            except ElementNotEnabled:
                sleep(1)
                _window.type_keys(command)
            sleep(step_delay)

    @staticmethod
    def close_warning():
        excel_pid = Utils.get_current_process_pid(proc_name='EXCEL.EXE')
        app = pywinauto.Application(backend='uia').connect(process=excel_pid)
        for win in app.windows():
            win_text = win.window_text()
            if not win_text:
                continue
            window = app.window(title=win_text)
            window['Закрыть'].click()

    @staticmethod
    def save_excel(file_path: str) -> None:
        Utils.close_warning()
        excel = win32.Dispatch('Excel.Application')
        excel.DisplayAlerts = False
        wb = excel.ActiveWorkbook
        wb.SaveAs(file_path, FileFormat=20)
        wb.Close(True)
        Utils.kill_all_processes(proc_name='EXCEL')

    @staticmethod
    def is_key_present(key: str, rows: List[Dict[str, str]]) -> bool:
        return next((True for row in rows if key in row), False)

    @staticmethod
    def is_kvit_required(rows: List[Dict[str, str]]) -> bool:
        return next((True for row in rows if row['KVITFL'] != '1'), False)
