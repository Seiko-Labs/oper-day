import os
import pywinauto
from pywinauto import Application, WindowSpecification
from pywinauto.timings import TimeoutError as TimingsTimeoutError
from time import sleep
import datetime
from datetime import datetime as dt
from utils import Utils, BackendManager
import win32com.client
from typing import List, Dict, Tuple, Any
from dataclasses import dataclass
from data_structures import DateInfo, RobotWorkTime
from bot_notification import TelegramNotifier
import copy


class Actions:
    def __init__(self, app: Application, today: DateInfo,
                 robot_time: RobotWorkTime, notifier: TelegramNotifier) -> None:
        self.app = app
        self.utils = Utils()
        self.is_kvit_required = False
        self.today = today
        self.robot_time = robot_time
        self.notifier = notifier
        if today.date != dt.now().date():
            self._change_day(today.date_str)

    def _choose_mode(self, mode: str) -> None:
        mode_win = self.app.window(title='Выбор режима')
        mode_win['Edit2'].wrapper_object().set_text(text=mode)
        mode_win['Edit2'].wrapper_object().type_keys('~')

    def _save_file(self, name: str, add_col: bool = False) -> None:
        file_win = self.app.window(title='Выберите файл для экспорта')
        file_win['Edit0'].wrapper_object().set_text(text=name)
        file_win.wrapper_object().send_keystrokes('~')
        confirm_win = self.app.window(title='Confirm Save As')
        if confirm_win.exists():
            confirm_win['&Yes'].wrapper_object().click()

        sort_win = self._get_window(title='Сортировка', wait_for='exists')
        if add_col:
            sort_win.type_keys(f'{19 * "{DOWN}"}{{SPACE}}')
        sort_win['OK'].wrapper_object().click()

    # def _wait_for_reg_finish(self, file_name: str, reg_num: str) -> None:
    #     task_win = self._get_window(title='Задания на обработку операционных периодов')
    #     while True:
    #         self.utils.kill_all_processes(proc_name='EXCEL')
    #         self._refresh(task_win)
    #         self.utils.type_keys(task_win, '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}~')
    #
    #         self._save_file(name=file_name)
    #
    #         sleep(2)
    #         file_path: str = rf'C:\Temp\{file_name}.xls'
    #         while not os.path.isfile(path=file_path):
    #             sleep(2)
    #         self.utils.kill_all_processes(proc_name='EXCEL')
    #
    #         if self.utils.is_reg_procedure_ready(file_name=file_path, reg_num=reg_num):
    #             break
    #
    #     print('SUCCESS')

    def _wait_for_reg_finish(self, _window, file_name: str, reg: str, main_branch_selected: bool = False, delay: int = 10) -> None:
        while True:
            self.utils.kill_all_processes(proc_name='EXCEL')
            self._refresh(_window)
            self.utils.type_keys(_window, '{RIGHT}{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}~')

            self._save_file(name=file_name)

            sleep(2)
            file_path: str = rf'C:\Temp\{file_name}.xls'
            while not os.path.isfile(path=file_path):
                sleep(2)
            self.utils.kill_all_processes(proc_name='EXCEL')

            if reg == '4' and main_branch_selected:
                main_branch_data = next(row for row in self.utils.text_to_dicts(file_path) if row['Код подразделения'] == '00')
                if main_branch_data['CUSTFL4'] != '0':
                    break
            elif reg == '2' and not main_branch_selected:
                data = [row for row in self.utils.text_to_dicts(file_path) if row['Код подразделения'] != '00']
                if len([row for row in data if row['CUSTFL2'] != '0']) == len(data):
                    break
            elif reg == '2' and main_branch_selected:
                main_branch_data = next(row for row in self.utils.text_to_dicts(file_path) if row['Код подразделения'] == '00')
                if main_branch_data['CUSTFL2'] != '0':
                    break

            sleep(delay)

        print('SUCCESS')

    def _get_kvit_rows(self, _window: WindowSpecification) -> List[Dict[str, str]]:
        self.utils.type_keys(_window, '{VK_SHIFT down}{VK_MENU}д{VK_SHIFT up}{UP}~', step_delay=.2)

        sort_win = self._get_window(title='Сортировка')
        self.utils.type_keys(sort_win, f'{15 * "{DOWN}"}{{SPACE}}', step_delay=.05)
        sort_win['OK'].wrapper_object().click()
        sleep(1)

        temp_file_path = r'C:\Temp\kvit.txt'
        self.utils.save_excel(file_path=temp_file_path)

        kvit_rows = self.utils.text_to_dicts(file_name=temp_file_path)

        if not self.utils.is_key_present(key='KVITFL', rows=kvit_rows):
            kvit_rows = self._get_kvit_rows(_window)
        return kvit_rows

    def _export_excel(self, file_name: str, is_excel_closed: bool = True, add_col: bool = False) -> List[Dict[str, str]]:
        temp_file_path = rf'C:\Temp\{file_name}.xls'
        self._save_file(name=file_name, add_col=add_col)
        sleep(1)
        if is_excel_closed:
            while not os.path.isfile(path=temp_file_path):
                sleep(2)
        self.utils.kill_all_processes(proc_name='EXCEL')

        return self.utils.text_to_dicts(file_name=temp_file_path)

    def _close_day(self, _window: WindowSpecification) -> None:
        self.utils.type_keys(_window, '{VK_SHIFT down}{VK_MENU}оо{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{DOWN}~')

        close_day_win = self._get_window(title='Закрытие операционного периода')
        date_checkbox = close_day_win['CheckBox3'].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        close_day_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        branch_checkbox = close_day_win['CheckBox2'].wrapper_object()
        if branch_checkbox.get_check_state() == 0:
            branch_checkbox.click()
        sleep(.5)
        close_day_win.close()
        # close_day_win['OK'].wrapper_object().click()

    def _refresh(self, _window: WindowSpecification) -> None:
        window_text = _window.window_text()
        keystrokes = ''
        if window_text == 'Состояние операционных периодов':
            keystrokes = '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{UP}{UP}~'
        elif window_text == 'Задания на обработку операционных периодов':
            keystrokes = '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{DOWN}~'
        self.utils.type_keys(_window, keystrokes)

    def _select_all_branches(self, _window, to_bottom: bool = False) -> None:
        self.utils.type_keys(_window, '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{UP}~')
        select_win = self._get_window(title='Выделение', wait_for='exists')

        down = 5 * '{PGDN}'
        if to_bottom:
            select_win['OK'].wrapper_object().click()
            self.utils.type_keys(_window, f'{{UP}}{{DOWN}}{{VK_SHIFT down}}{down}{{VK_SHIFT up}}', step_delay=.01)
        else:
            select_win['Нижний уровень'].wrapper_object().click()
            select_win['OK'].wrapper_object().click()

            self._refresh(_window=_window)
            self.utils.type_keys(_window, f'{{LEFT}}{{LEFT}}{{VK_SHIFT down}}{down}{{VK_SHIFT up}}')

    def _reset_to_00(self, _window: WindowSpecification) -> None:
        self.utils.type_keys(_window, '{PGUP}{PGUP}{PGUP}{PGUP}{PGUP}', step_delay=.01)

    def _get_window(self, title: str, app: Application or None = None, wait_for: str = 'exists', timeout: int = 20,
                    regex: bool = False) -> WindowSpecification:
        if not app:
            app = self.app
        _window = app.window(title=title) if not regex else app.window(title_re=title)
        _window.wait(wait_for=wait_for, timeout=timeout)
        return _window

    def _change_day(self, _date: str, repeat: bool = False) -> None:
        with BackendManager(self.app, 'uia'):
            status_win = self._get_window(app=self.app, title='Банковская система.+', regex=True)
            if status_win['Static3'].window_text().strip() == _date:
                return
            status_win['Static3'].double_click_input()

            oper_day_win = self._get_window(app=self.app, title='Текущий операционный день')
            oper_day_win['Edit2'].set_text(text=_date)
            oper_day_win['OK'].click()

            if not repeat:
                warning_win = self._get_window(app=self.app, title='Внимание')
                warning_win['Да'].click()

            oper_day_win.Dialog.wait(wait_for='exists', timeout=20)
            oper_day_win.Dialog.type_keys('~')

    def _press_status_button(self, _window: WindowSpecification, button: str = 'Все задания на обработку', pixel_step: int = 30) -> None:
        status_win = self.app.window(title_re='Банковская система.+')
        mid_point = _window['Static0'].rectangle().mid_point()
        i, x, y = 0, mid_point.x, mid_point.y
        while status_win['StatusBar'].window_text() != button:
            x, y = mid_point.x - i * pixel_step, mid_point.y
            _window.move_mouse_input(coords=(x, y), absolute=True)
            i += 1
        sleep(.1)
        _window.click_input(button='left', coords=(x, y), absolute=True)

    def step1(self) -> None:
        # выбор режима COPPER
        self._choose_mode(mode='COPPER')

        # окно Состояние операционных периодов
        main_win = self._get_window(title='Состояние операционных периодов')

        # Снять признак выполнения 4
        self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}р{VK_SHIFT up}{UP}~')

        # Подтверждение "снятие признака выполнения 4"
        try:
            confirm_win = self._get_window(title='Подтверждение', timeout=5)
            confirm_win['Да'].wrapper_object().click()
        except TimingsTimeoutError:
            pass

        # Регламентарная процедура 4
        self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}р{VK_SHIFT up}{UP}{UP}~')

        procedure_win = self._get_window(title='Регламентная процедура 4')

        date_checkbox = procedure_win['CheckBox3'].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        procedure_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        branch_checkbox = procedure_win['CheckBox2'].wrapper_object()
        if branch_checkbox.get_check_state() == 1:
            branch_checkbox.click()
        procedure_win['OK'].wrapper_object().click()

        confirm_win = self._get_window(title='Подтверждение')
        confirm_win['Да'].wrapper_object().click()

        self._wait_for_reg_finish(_window=main_win, file_name='reg_procedure_4', reg='4', main_branch_selected=True)

    def step2(self) -> None:
        self._choose_mode(mode='COPPER')

        main_win = self._get_window(title='Состояние операционных периодов')

        self._select_all_branches(_window=main_win, to_bottom=True)

        # Регламентарная процедура 2
        self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}р{VK_SHIFT up}{DOWN}~')
        procedure_win = self._get_window(title='Регламентная процедура 2')

        date_checkbox = procedure_win['CheckBox2'].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        procedure_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        procedure_win['OK'].wrapper_object().click()

        confirm_win = self._get_window(title='Подтверждение')
        confirm_win['Да'].wrapper_object().click()

        self._wait_for_reg_finish(_window=main_win, file_name='reg_procedure_2', reg='2')
        self._select_all_branches(_window=main_win)

    def step3(self) -> None:
        self._choose_mode(mode='COPPER')

        main_win = self.app.window(title='Состояние операционных периодов')

        self._select_all_branches(_window=main_win)
        self._reset_to_00(_window=main_win)

        # Регламентарная процедура 2
        self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}р{VK_SHIFT up}{DOWN}~')
        procedure_win = self._get_window(title='Регламентная процедура 2')

        date_checkbox = procedure_win['CheckBox3'].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        procedure_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        branch_checkbox = procedure_win['CheckBox2'].wrapper_object()
        if branch_checkbox.get_check_state() == 1:
            branch_checkbox.click()
        procedure_win['OK'].wrapper_object().click()

        confirm_win = self._get_window(title='Подтверждение')
        confirm_win['Да'].wrapper_object().click()

        self._wait_for_reg_finish(_window=main_win, file_name='reg_procedure_2', reg='2', main_branch_selected=True)

    def step4(self) -> None:
        self._choose_mode(mode='EXTRCT')

        filter_win = self._get_window(title='Фильтр')

        filter_win['Edit8'].wrapper_object().set_text(text=self.today.date_str)
        filter_win['OKButton'].wrapper_object().click()

        main_win = self._get_window(title='Выписки')

        self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}~')
        rows = self._export_excel(file_name='950rows')

        try:
            down = next((i for i, x in enumerate(rows) if x['Тип'] == '950'), None) * '{DOWN}'
        except KeyError:
            self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}~')
            rows = self._export_excel(file_name='950rows', add_col=True)
            down = next((i for i, x in enumerate(rows) if x['Тип'] == '950'), None) * '{DOWN}'
        self.utils.type_keys(main_win, f'{down}~')

        vypiska_win = self._get_window(title='Выписка', wait_for='exists active')
        sleep(.1)
        vypiska_win.type_keys('~')

        document_win = self._get_window(title='Документ')
        kvit_rows = self._get_kvit_rows(_window=document_win)
        print(kvit_rows)

        self.is_kvit_required = self.utils.is_kvit_required(rows=kvit_rows)

        document_win.close()
        vypiska_win.close()
        main_win.close()

    def step5(self) -> None:
        # if not self.is_kvit_required:
        #     return

        self._choose_mode(mode='AUTCHK')

        main_win = self._get_window(title='Выверяемые счета')

        self.utils.type_keys(main_win, '{F7}')

        search_win = self._get_window(title='Поиск по номеру счета')
        search_win['Edit2'].wrapper_object().set_text(text='KZ139')
        search_win['OK'].wrapper_object().click()

        self.utils.type_keys(main_win, '^g')

        vyverka_win = self._get_window(title='Подготовка системы выверки')
        vyverka_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)

        self.utils.type_keys(main_win, '^p')

        info_win = self._get_window(title='Информация')
        info_win['OK'].wrapper_object().click()

        result_win = self._get_window(title='Результаты операции')
        result_text = result_win['Edit'].wrapper_object().window_text().strip()
        result_win.close()

        if 'Не успешно' in result_text:
            self.utils.type_keys(main_win, '^a')
            result_win = self._get_window(title='Результаты операции')
            result_win.close()

    def step6(self) -> None:
        self._choose_mode(mode='SORDPAY')

        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit8'].wrapper_object().set_text(self.today.date_str)
        filter_win['Edit6'].wrapper_object().set_text(self.today.date_str)
        filter_win['OK'].wrapper_object().click()

        main_win = self._get_window(title='Расчетные документы филиала', timeout=600)
        self.utils.type_keys(main_win, '{F12}')

        template_win = self._get_window(title='Шаблоны платежей')
        self.utils.type_keys(template_win, '{F9}')

        filter_win2 = self._get_window(title='Фильтр')
        filter_win2['Edit2'].wrapper_object().set_text(text='Вечер')
        filter_win2['OK'].wrapper_object().click()

        pass

    def step7(self) -> None:
        self._choose_mode(mode='COPPER')

        main_win = self._get_window(title='Состояние операционных периодов')

        self._select_all_branches(_window=main_win)
        self._close_day(_window=main_win)
        self._reset_to_00(_window=main_win)
        self._close_day(_window=main_win)
        self._change_day(_date=self.today.next_date_str)
        self._refresh(_window=main_win)
        self._change_day(_date=self.today.date_str, repeat=True)
        self._refresh(_window=main_win)
        pass

    def step8(self) -> None:
        self._choose_mode(mode='SORDPAY')

        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit8'].wrapper_object().set_text(self.today.date_str)
        filter_win['Edit6'].wrapper_object().set_text(self.today.date_str)

        # main_win = self._get_window(title='Состояние операционных периодов')

        pass

    def _reg1(self, _window, all_branches):
        self.utils.type_keys(_window, '{VK_SHIFT down}{VK_MENU}р{VK_SHIFT up}~')
        procedure_win = self._get_window(title='Регламентная процедура 1')

        date_checkbox = procedure_win['CheckBox3'].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        procedure_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        branch_checkbox = procedure_win['CheckBox2'].wrapper_object()
        if branch_checkbox.get_check_state() == all_branches:
            branch_checkbox.click()
        procedure_win['OK'].wrapper_object().click()

        confirm_win = self._get_window(title='Подтверждение')
        confirm_win['НЕТ'].wrapper_object().click()
        # confirm_win['ДА'].wrapper_object().click()

    def step9(self) -> None:
        self._choose_mode(mode='COPPER')

        main_win = self._get_window(title='Состояние операционных периодов')
        self._select_all_branches(_window=main_win)
        # Регламентная процедура 1
        self._reg1(_window=main_win, all_branches=1)

        main_win.wait(wait_for='visible', timeout=20)

        # WAIT FOR END

        self._reset_to_00(_window=main_win)
        self._refresh(_window=main_win)
        self._reg1(_window=main_win, all_branches=0)

        pass

    def step10(self):
        self.notifier.send_notification('step10')
        self._choose_mode(mode='COPPER')

        main_win = self._get_window(title='Состояние операционных периодов')
        self.utils.type_keys(_window=main_win, keystrokes='{F5}')

        report_win = self._get_window(title='Выбор отчета')
        self.utils.type_keys(_window=report_win, keystrokes='{F9}')

        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit4'].wrapper_object().set_text(text='PC05_101')
        filter_win['OK'].wrapper_object().click()

        report_win.wait(wait_for='visible active exists', timeout=20)
        report_win['Предварительный просмотр'].wrapper_object().click()

        report_win['Экспорт в файл...'].wrapper_object().click()

        file_win = self._get_window(title='Файл отчета ')
        try:
            file_win['ComboBox'].wrapper_object().select(2)
        except (IndexError, ValueError):
            pass
        file_win['OK'].wrapper_object().click()

        params_win = self._get_window(title='Параметры отчета ')
        params_win['Edit2'].wrapper_object().set_text(text=self.robot_time.start_str)
        self.robot_time.update()
        params_win['Edit4'].wrapper_object().set_text(text=self.robot_time.end_str)
        # params_win['OK'].wrapper_object().click()
        self.notifier.send_notification('end step10')
        pass

    def run(self) -> None:
        # method_list = [func for func in dir(self) if callable(getattr(self, func)) and 'step' in func]
        # for method in method_list:
        #     getattr(self, method)()
        # self.step1()
        # self.step2()
        self.step3()
        # self.step4()
        # self.step5()
        # self.step6()
        # self.step7()
        # self.step8()
        # self.step9()
        # self.step10()
