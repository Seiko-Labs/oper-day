import datetime
import os
from datetime import datetime as dt
from time import sleep
from typing import List, Dict
import psutil
import pywinauto
import win32com.client as win32


class Utils:
    def __init__(self):
        self.file_types = {
            'xlsb': 50,
            'xlsx': 51
        }

    def convert(self, src_file: str, dst_file: str, file_type: str):
        file_format = self.file_types[file_type]
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(src_file)
        wb.SaveAs(dst_file, FileFormat=file_format)
        wb.Close()
        excel.Application.Quit()

    @staticmethod
    def kill_process(pid):
        for proc in psutil.process_iter():
            try:
                if pid != proc.pid:
                    continue
                p: psutil.Process = psutil.Process(proc.pid)
                p.terminate()
            except psutil.AccessDenied:
                continue

    @staticmethod
    def get_excel_pids():
        excel_pids = []
        for proc in psutil.process_iter():
            try:
                if 'EXCEL' not in proc.name():
                    continue
                pid = proc.pid
                _: psutil.Process = psutil.Process(proc.pid)
                excel_pids.append(pid)
            except psutil.AccessDenied:
                continue
        return excel_pids

    @staticmethod
    def kill_all_excel() -> None:
        for proc in psutil.process_iter():
            if 'EXCEL' not in proc.name():
                continue
            try:
                p: psutil.Process = psutil.Process(proc.pid)
                p.terminate()
            except psutil.AccessDenied:
                continue

    # def get_offset(self) -> int:
    #     xlsx_file_path: str = self.convert_to_xlsx()
    #     wb: Workbook = openpyxl.load_workbook(filename=xlsx_file_path)
    #     sh: Worksheet = wb.active
    #     rows = [r for r in sh.iter_rows(min_row=2, max_row=sh.max_row, values_only=True) if
    #             not any('записи' in str(c) for c in r)]
    #     self.kill_all_excel()
    #     os.unlink(xlsx_file_path)
    #     return next(i for i, r in enumerate(rows[1::]) for c in r if c == 950)

    @staticmethod
    def get_current_colvir_pid() -> int:
        res: int or None = None
        for proc in psutil.process_iter():
            if 'COLVIR' in proc.name():
                res = proc.pid
        return res


def is_ready(file_name: str, current_date: datetime.datetime, delay: int = 5) -> bool:
    with open(file=file_name, mode='r', encoding='utf-16') as f:
        rows: List[List[str]] = [line.split('\t') for line in f.readlines()]
    if len(rows) == 3:
        sleep(delay)
        return False
    keys: List[str] = list(rows[1])
    data: List[Dict] = []
    for row in rows[2:-1]:
        data.append(dict(zip(keys, row)))

    for x in data:
        if (current_date - dt.strptime(x['Дата уст.'], '%d.%m.%y %H:%M:%S')).total_seconds() > 30:
            continue
        if x['Состояние'] != 'Не обработано':
            os.unlink(file_name)
            return True
    sleep(delay)
    return False


def main():
    utils = Utils()
    app = pywinauto.Application(backend='win32').connect(process=5800)

    task_win = app.window(title='Задания на обработку операционных периодов')
    task_win.wait(wait_for='exists', timeout=20)

    while True:
        utils.kill_all_excel()
        task_win.wrapper_object().menu_item('Список').sub_menu().items()[6].select()
        task_win.wrapper_object().menu_item('Список').sub_menu().items()[4].sub_menu().items()[1].select()

        file_win = app.window(title='Выберите файл для экспорта')
        temp_file_name = 'test'
        file_win['Edit0'].wrapper_object().set_text(text=temp_file_name)
        file_win.wrapper_object().send_keystrokes('{ENTER}')
        confirm_win = app.window(title='Confirm Save As')
        if confirm_win.exists():
            confirm_win['&Yes'].wrapper_object().click()

        sort_win = app.window(title='Сортировка')
        sort_win.wait(wait_for='exists', timeout=5)
        sort_win['OK'].wrapper_object().click()
        current_date = dt.now()

        file_path = rf'C:\Temp\{temp_file_name}.xls'
        while not os.path.isfile(file_path):
            sleep(1)

        if is_ready(file_name=file_path, current_date=current_date):
            break

    print('SUCCESS')


if __name__ == '__main__':
    main()
