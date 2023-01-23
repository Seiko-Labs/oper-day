import psutil
from psutil import Process
import pywinauto
import win32com.client as win32
from typing import List, Dict
import re
from itertools import islice
from excel_converter import ExcelConverter
from time import sleep
from datetime import datetime as dt


class Utils:
    def __init__(self) -> None:
        self.excel_converter: ExcelConverter = ExcelConverter()

    def convert(self, src_file: str, dst_file: str, file_type: str) -> None:
        self.excel_converter.convert(src_file=src_file, dst_file=dst_file, file_type=file_type)

    @staticmethod
    def kill_process(pid) -> None:
        p: Process = Process(pid)
        p.terminate()

    @staticmethod
    def kill_all_processes(proc_name: str, restricted_pids: List[int] or None = None) -> None:
        processes_to_kill: List[Process] = [Process(proc.pid) for proc in psutil.process_iter() if proc_name in proc.name()]
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
    def text_to_dicts(file_name: str) -> List[Dict[str, str]]:
        pattern = re.compile(r'(Начало|Конец) записи \d+\.\d+\.\d+ \d+:\d+:\d+')
        encoding = 'utf-8' if file_name.endswith('.txt') else 'utf-16'
        with open(file=file_name, mode='r', encoding=encoding) as file:
            rows = [[el.replace('\n', '') for el in line.split('\t')] for line in file if not pattern.search(line)]
        header = [col.strip() for col in rows[0]]
        data_rows = islice(rows, 1, None)
        return [{col: val.strip() for col, val in zip(header, row)} for row in data_rows]

    @staticmethod
    def is_reg_4_ready(file_name: str, current_date: dt, delay: int = 5) -> bool:
        data = Utils.text_to_dicts(file_name=file_name)
        if not data:
            sleep(delay)
            return False

        for x in data:
            if abs(current_date - dt.strptime(x['Дата уст.'], '%d.%m.%y %H:%M:%S')).total_seconds() > 30:
                continue
            if x['Состояние'] != 'Не обработано':
                return True
        sleep(delay)
        return False

    @staticmethod
    def type_keys(window, keystrokes: str, step_delay: float = .1) -> None:
        for command in list(filter(None, re.split(r'({.+?})', keystrokes))):
            try:
                window.type_keys(command)
            except pywinauto.base_wrapper.ElementNotEnabled:
                sleep(1)
                window.type_keys(command)
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
