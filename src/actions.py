import os
import re
from dataclasses import fields
from datetime import datetime as dt
from time import sleep
from typing import List, Dict
from pywinauto import Application, WindowSpecification
from pywinauto.base_wrapper import ElementNotEnabled
from pywinauto.timings import TimeoutError as TimingsTimeoutError
from data_structures import Notifiers, DateInfo, RobotWorkTime, Buttons, Button
from utils import Utils, BackendManager


class Actions:
    def __init__(self, app: Application, today: DateInfo,
                 robot_time: RobotWorkTime, notifiers: Notifiers) -> None:
        self.app = app
        self.utils = Utils()
        self.is_kvit_required = False
        self.today = today
        self.robot_time = robot_time
        self.notifiers = notifiers
        self.buttons: Buttons = Buttons()
        if today.date != dt.now().date():
            self._change_day(today.date_str)

    def _get_buttons(self, main_win: WindowSpecification, pixel_step: int = 5, offset: int = 1):
        status_win = self.app.window(title_re='Банковская система.+')
        rectangle = main_win['Static0'].rectangle()
        mid_point = rectangle.mid_point()
        main_win.move_mouse_input(coords=(mid_point.x, mid_point.y), absolute=True)
        left_border = rectangle.left
        i, x, y = 0, left_border, mid_point.y

        while self.buttons.filled_count < 8:
            x, y = left_border + i * pixel_step, mid_point.y
            main_win.move_mouse_input(coords=(x, y), absolute=True)
            i += 1
            button_name = status_win['StatusBar'].window_text().strip()
            if button_name == 'Для помощи нажмите F1':
                continue
            for field in fields(self.buttons):
                button = getattr(self.buttons, field.name)
                if field.name != 'tasks_refresh' and isinstance(button,
                                                                Button) and button_name == button.name and button.filled is False:
                    button.coords = (x + offset, y)
                    button.filled = True
                    self.buttons.filled_count += 1

        main_win.click_input(coords=self.buttons.tasks.coords, absolute=True)
        tasks_win = self.app.window(title='Задания на обработку операционных периодов')

        main_win.move_mouse_input(coords=(mid_point.x, mid_point.y), absolute=True)
        left_border = rectangle.left
        i, x, y = 0, left_border, mid_point.y

        while self.buttons.tasks_refresh.coords == (0, 0):
            x, y = left_border + i * pixel_step, mid_point.y
            tasks_win.move_mouse_input(coords=(x, y), absolute=True)
            i += 1
            button_name = status_win['StatusBar'].window_text().strip()
            if button_name == 'Обновить список':
                self.buttons.tasks_refresh.coords = (x + offset, y)
                self.buttons.tasks_refresh.filled = True
                self.buttons.filled_count += 1
        tasks_win.close()

        self._choose_mode(mode='SORDPAY')
        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit8'].wrapper_object().set_text(self.today.date_str)
        filter_win['Edit6'].wrapper_object().set_text(self.today.date_str)
        filter_win['Edit2'].wrapper_object().set_text('1')
        filter_win['Edit4'].wrapper_object().set_text('1')
        sleep(1)
        filter_win['OK'].wrapper_object().click()

        main_win = self._get_window(title='Расчетные документы филиала', timeout=600)
        self.utils.type_keys(main_win, '{F12}')

        template_win = self._get_window(title='Шаблоны платежей')
        self.utils.type_keys(template_win, '{F9}')

        filter_win2 = self._get_window(title='Фильтр')
        filter_win2['Edit2'].wrapper_object().type_keys('Вечер~~', pause=.2)

        template_win = self.app.window(title='Шаблоны платежей')
        sleep(1)
        if template_win.exists():
            template_win['OK'].wrapper_object().click()

        order_win = self._get_window(title='Мемориальный ордер', timeout=360)

        rectangle = order_win['Static0'].rectangle()
        mid_point = rectangle.mid_point()
        order_win.move_mouse_input(coords=(mid_point.x, mid_point.y), absolute=True)
        left_border = rectangle.left
        i, x, y = 0, left_border, mid_point.y

        while self.buttons.filled_count < 11:
            x, y = left_border + i * pixel_step, mid_point.y
            order_win.move_mouse_input(coords=(x, y), absolute=True)
            i += 1
            button_name = status_win['StatusBar'].window_text().strip()
            if button_name == 'Сохранить изменения (PgDn)':
                self.buttons.save.coords = (x + offset, y)
                self.buttons.filled_count += 1
            elif button_name == 'Выполнить операцию':
                self.buttons.operations.coords = (x + offset, y)
                self.buttons.filled_count += 1
        order_win.close()
        main_win.close()

    def _choose_mode(self, mode: str) -> None:
        mode_win = self.app.window(title='Выбор режима')
        mode_win['Edit2'].wrapper_object().set_text(text=mode)
        mode_win['Edit2'].wrapper_object().type_keys('~')

    def _save_file(self, name: str, add_col: bool = False) -> None:
        if os.path.exists(fr'C:\Temp\{name}.xls'):
            os.unlink(fr'C:\Temp\{name}.xls')

        file_win = self.app.window(title='Выберите файл для экспорта')
        file_win['Edit0'].wrapper_object().click_input()
        file_win['Edit0'].wrapper_object().type_keys(f'{name}~')

        sort_win = self._get_window(title='Сортировка', wait_for='exists')
        if add_col:
            sort_win.type_keys(f'{19 * "{DOWN}"}{{SPACE}}')
        sort_win['OK'].wrapper_object().click()

    def _wait_for_reg_finish(self, main_win, file_name: str, delay: int = 60) -> None:
        finished = False
        while not finished:
            self.notifiers.log.send_message(message=f'Ожидание окончания обработки процедуры...')
            self.utils.kill_all_processes(proc_name='EXCEL')
            main_win.click_input(button='left', coords=self.buttons.tasks_refresh.coords, absolute=True)
            self.utils.type_keys(main_win, '{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}~')

            self._save_file(name=file_name)

            sleep(2)
            file_path: str = rf'C:\Temp\{file_name}.xls'
            while not os.path.isfile(path=file_path):
                sleep(2)
            self.utils.kill_all_processes(proc_name='EXCEL')

            finished = not [row for row in self.utils.text_to_dicts(file_path)
                            if row['Исполнитель'] == 'Блокировка учетных записей']

            sleep(delay)

        self.notifiers.log.send_message(message=f'Процедура успешно обработана')

    def _get_kvit_rows(self, _window: WindowSpecification) -> List[Dict[str, str]]:
        self.utils.type_keys(_window, '{VK_SHIFT down}{VK_MENU}д{VK_SHIFT up}{UP}~', step_delay=.2)

        sort_win = self._get_window(title='Сортировка')
        self.utils.type_keys(sort_win, f'{15 * "{DOWN}"}{{SPACE}}', step_delay=.05)
        sort_win['OK'].wrapper_object().click()
        sleep(1)

        temp_file_path = r'C:\Temp\kvit.txt'
        self.utils.save_excel(file_path=temp_file_path)

        kvit_rows = self.utils.text_to_dicts(file_path=temp_file_path)

        if not self.utils.is_key_present(key='KVITFL', rows=kvit_rows):
            kvit_rows = self._get_kvit_rows(_window)
        return kvit_rows

    def _export_excel(self, file_name: str, is_excel_closed: bool = True,
                      add_col: bool = False) -> List[Dict[str, str]]:
        temp_file_path = rf'C:\Temp\{file_name}.xls'
        self._save_file(name=file_name, add_col=add_col)
        sleep(1)
        if is_excel_closed:
            while not os.path.isfile(path=temp_file_path):
                sleep(2)
        self.utils.kill_all_processes(proc_name='EXCEL')

        return self.utils.text_to_dicts(file_path=temp_file_path)

    def _refresh(self, _window: WindowSpecification) -> None:
        _window.click_input(button='left', coords=self.buttons.refresh.coords, absolute=True)

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

    def _reset_to_00(self, main_win: WindowSpecification) -> None:
        self.utils.type_keys(main_win, '{PGUP}{PGUP}{PGUP}{PGUP}{PGUP}{RIGHT}{DOWN}{LEFT}', step_delay=.01)

    def _get_window(self, title: str, app: Application or None = None, wait_for: str = 'exists', timeout: int = 20,
                    regex: bool = False, found_index: int = 0) -> WindowSpecification:
        if not app:
            app = self.app
        _window = app.window(title=title, found_index=found_index)\
            if not regex else app.window(title_re=title, found_index=found_index)
        _window.wait(wait_for=wait_for, timeout=timeout)
        sleep(.5)
        return _window

    def _change_day(self, _date: str) -> None:
        with BackendManager(self.app, 'uia'):
            status_win = self._get_window(app=self.app, title='Банковская система.+', regex=True)
            if status_win['Static3'].window_text().strip() == _date:
                return
            status_win['Static3'].double_click_input()

            oper_day_win = self._get_window(app=self.app, title='Текущий операционный день')
            oper_day_win['Edit2'].set_text(text=_date)
            oper_day_win['OK'].click()

            try:
                warning_win = self._get_window(app=self.app, title='Внимание', timeout=5)
                warning_win['Да'].click()
            except TimingsTimeoutError:
                pass

            oper_day_win.Dialog.wait(wait_for='exists', timeout=20)
            oper_day_win.Dialog.type_keys('~')

    def _reset_tasks(self):
        self._choose_mode(mode='SYST')
        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit6'].wrapper_object().set_text('%021%')
        filter_win['OK'].click()

        main_win = self._get_window(title='Фоновые задания')
        self.utils.type_keys(main_win, f'{"{VK_CONTROL down}{F10}{VK_CONTROL up}{DOWN}" * 31}{"{F10}{UP}" * 31}')
        main_win.close()

    def _fill_procedure_form(self, procedure_win: WindowSpecification, main_win: WindowSpecification,
                             main_branch_selected: bool, file_name: str) -> None:
        checkbox = 'CheckBox3' if main_branch_selected else 'CheckBox2'
        date_checkbox = procedure_win[checkbox].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            sleep(1)
            date_checkbox.click_input()
        try:
            procedure_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        except ElementNotEnabled:
            procedure_win['CheckBox3'].wrapper_object().click_input()
            procedure_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        self.notifiers.log.send_message(message=f'Введен день {self.today.date_str} в форму процедуры')

        if main_branch_selected:
            branch_checkbox = procedure_win['CheckBox2'].wrapper_object()
            if branch_checkbox.get_check_state() == 1:
                branch_checkbox.click()
            self.notifiers.log.send_message(message=f'Убрана галочка с "Все филиалы" в форме процедуры')

        procedure_win['OK'].wrapper_object().click()
        confirm_win = self._get_window(title='Подтверждение')
        confirm_win.type_keys('~')
        self.notifiers.log.send_message(message=f'Регламентная процедура начата')

        main_win.click_input(button='left', coords=self.buttons.tasks.coords, absolute=True)
        tasks_win = self._get_window(title='Задания на обработку операционных периодов')

        self._wait_for_reg_finish(
            main_win=tasks_win,
            file_name=file_name,
        )

    def _close_day(self, main_win: WindowSpecification, main_branch_selected: bool = False) -> None:
        main_win.click_input(button='left', coords=self.buttons.close_oper_day.coords, absolute=True)

        close_day_win = self._get_window(title='Закрытие операционного периода')
        date_checkbox = close_day_win['CheckBox3'].wrapper_object()
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        try:
            close_day_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        except ElementNotEnabled:
            close_day_win['CheckBox3'].wrapper_object().click()
            close_day_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        branch_checkbox = close_day_win['CheckBox2'].wrapper_object()
        if branch_checkbox.get_check_state() == (1 if main_branch_selected else 0):
            branch_checkbox.click()
        sleep(.5)
        close_day_win['OK'].wrapper_object().click()
        try:
            confirm_win = self._get_window(title='Подтверждение', timeout=5)
            confirm_win.type_keys('~')
        except TimingsTimeoutError:
            pass

    def _open_day(self, main_win: WindowSpecification, main_branch_selected: bool = False) -> None:
        main_win.click_input(button='left', coords=self.buttons.close_oper_day.coords, absolute=True)

        open_day_win = self._get_window(title='Новый операционный период')
        date_checkbox = open_day_win['CheckBox3'].wrapper_object()
        sleep(1)
        if date_checkbox.get_check_state() == 0:
            date_checkbox.click()
        try:
            open_day_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        except ElementNotEnabled:
            open_day_win['CheckBox3'].wrapper_object().click()
            open_day_win['Edit2'].wrapper_object().set_text(text=self.today.next_work_date_str)
        branch_checkbox = open_day_win['CheckBox2'].wrapper_object()
        if branch_checkbox.get_check_state() == (1 if main_branch_selected else 0):
            branch_checkbox.click()
        sleep(.5)
        open_day_win['OK'].wrapper_object().click()
        try:
            confirm_win = self._get_window(title='Подтверждение', timeout=5)
            confirm_win.type_keys('~')
        except TimingsTimeoutError:
            pass

    def step1(self) -> None:
        """DONE"""

        main_win = self._get_window(title='Состояние операционных периодов')

        self.notifiers.log.send_message(message='Снятие признака выполнения 4 в режиме COPPER')
        main_win.click_input(button='left', coords=self.buttons.remove_reg_procedure_4.coords, absolute=True)

        try:
            confirm_win = self._get_window(title='Подтверждение', timeout=5)
            confirm_win.type_keys('~')
            self.notifiers.log.send_message(message='Признак выполнения 4 снят')
        except TimingsTimeoutError:
            pass

        main_win.click_input(button='left', coords=self.buttons.reg_procedure_4.coords, absolute=True)

        procedure_win = self._get_window(title='Регламентная процедура 4', found_index=0)
        self.notifiers.log.send_message(message='Регламентная процедура 4 по 00 в режиме COPPER')

        self._fill_procedure_form(
            procedure_win=procedure_win,
            main_win=main_win,
            main_branch_selected=True,
            file_name='reg_procedure_4',
        )

    def step2(self) -> None:
        main_win = self._get_window(title='Состояние операционных периодов')

        self._select_all_branches(_window=main_win, to_bottom=True)

        self.notifiers.log.send_message(message=f'Регламентная процедура 2 в режиме COPPER')
        main_win.click_input(button='left', coords=self.buttons.reg_procedure_2.coords, absolute=True)
        procedure_win = self._get_window(title='Регламентная процедура 2')

        self._fill_procedure_form(
            procedure_win=procedure_win,
            main_win=main_win,
            main_branch_selected=False,
            file_name='reg_procedure_2_all',
        )

    def step3(self) -> None:
        main_win = self.app.window(title='Состояние операционных периодов')

        self._select_all_branches(_window=main_win)
        self._reset_to_00(main_win=main_win)

        main_win.click_input(button='left', coords=self.buttons.reg_procedure_2.coords, absolute=True)
        procedure_win = self._get_window(title='Регламентная процедура 2')
        self.notifiers.log.send_message(message=f'Регламентная процедура 2 по 00 в режиме COPPER')

        self._fill_procedure_form(
            procedure_win=procedure_win,
            main_win=main_win,
            main_branch_selected=True,
            file_name='reg_procedure_2_00',
        )
        main_win.close()

    def step4(self) -> None:
        self._choose_mode(mode='EXTRCT')

        filter_win = self._get_window(title='Фильтр')

        filter_win['Edit8'].wrapper_object().set_text(text=self.today.date_str)
        filter_win['OKButton'].wrapper_object().click()

        self.notifiers.log.send_message(message=f'Начало работы с выписками в режиме EXTRCT')

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

        self.is_kvit_required = self.utils.is_kvit_required(rows=kvit_rows)

        document_win.close()
        vypiska_win.close()
        main_win.close()

    def step5(self) -> None:
        if not self.is_kvit_required:
            self.notifiers.log.send_message(message=f'Квитовка выписок не требуется')
            return
        self.notifiers.log.send_message(message=f'Требуется квитовка выписок в режиме AUTCHK')

        self._choose_mode(mode='AUTCHK')

        main_win = self._get_window(title='Выверяемые счета')

        self.utils.type_keys(main_win, '{F7}')

        search_win = self._get_window(title='Поиск по номеру счета')
        search_win['Edit2'].wrapper_object().set_text(text='KZ139')
        search_win['OK'].wrapper_object().click()

        self.utils.type_keys(main_win, '^g')

        vyverka_win = self._get_window(title='Подготовка системы выверки')
        vyverka_win['Edit2'].wrapper_object().set_text(text=self.today.date_str)
        vyverka_win['OK'].wrapper_object().click()

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
        main_win.close()
        self.notifiers.log.send_message(message=f'Успешная квитовка выписок')

    def step6(self) -> None:
        self._choose_mode(mode='SORDPAY')
        self.notifiers.log.send_message(message=f'Начало работы в режиме SORDPAY')

        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit8'].wrapper_object().set_text(self.today.date_str)
        filter_win['Edit6'].wrapper_object().set_text(self.today.date_str)
        filter_win['Edit2'].wrapper_object().set_text('1')
        filter_win['Edit4'].wrapper_object().set_text('1')
        sleep(1)
        filter_win['OK'].wrapper_object().click()

        main_win = self._get_window(title='Расчетные документы филиала', timeout=600)
        self.utils.type_keys(main_win, '{F12}')

        template_win = self._get_window(title='Шаблоны платежей')
        self.utils.type_keys(template_win, '{F9}')

        filter_win2 = self._get_window(title='Фильтр')
        filter_win2['Edit2'].wrapper_object().type_keys('Вечер~~', pause=.2)
        self.notifiers.log.send_message(message=f'Выбран шаблон "Вечер"')

        template_win = self.app.window(title='Шаблоны платежей')
        sleep(1)
        if template_win.exists():
            template_win['OK'].wrapper_object().click()

        order_win = self._get_window(title='Мемориальный ордер', timeout=360)

        remainder_sum = ''.join(re.findall(r'[\d.]', order_win['Edit4'].wrapper_object().window_text().strip()))
        self.notifiers.log.send_message(message=f'Перенесена сумма {remainder_sum} в поле "Сумма"')
        order_win['Edit26'].type_keys(f'{remainder_sum}~', pause=.1)
        sleep(2)
        order_win.click_input(button='left', coords=self.buttons.save.coords, absolute=True)
        self.notifiers.log.send_message(message=f'Документ "Вечер" сохранен')
        sleep(4)

        order_win.close()
        sleep(2)

        main_win.click_input(button='left', coords=self.buttons.operations.coords, absolute=True)
        main_win.type_keys('{DOWN}~', pause=.1)
        self.notifiers.log.send_message(message=f'Сумма оплачена по документу "Вечер"')

        sleep(2)
        main_win.close()

    def _wait_for_day_procedure_end(self, main_win, file_name: str, procedure_type: str,
                                    main_branch_selected: bool = False, delay: int = 10) -> None:
        while True:
            self.utils.kill_all_processes(proc_name='EXCEL')
            self._refresh(main_win)
            self.utils.type_keys(main_win,
                                 '{RIGHT}{VK_SHIFT down}{VK_MENU}с{VK_SHIFT up}{DOWN}{DOWN}{DOWN}{RIGHT}{DOWN}~')

            self._save_file(name=file_name)

            sleep(2)
            file_path: str = rf'C:\Temp\{file_name}.xls'
            while not os.path.isfile(path=file_path):
                sleep(2)
            self.utils.kill_all_processes(proc_name='EXCEL')

            rows = self.utils.text_to_dicts(file_path)

            if procedure_type == 'close':
                date_key = 'Закрыты'
                date = self.today.date_str
            else:
                date_key = 'Открыты'
                date = self.today.next_work_date_str

            if main_branch_selected:
                data = [row for row in rows if
                        row['Код подразделения'] == '00' and row[f'{date_key} по дату'] == date]
            else:
                data = [row for row in rows
                        if row['Код подразделения'] != '00' and row[f'{date_key} по дату'] == date]

            if not data:
                break

            sleep(delay)

        print('SUCCESS')

    def step7(self) -> None:
        self._choose_mode(mode='COPPER')

        main_win = self._get_window(title='Состояние операционных периодов')

        self._select_all_branches(_window=main_win)

        self.notifiers.log.send_message(message=f'Начало закрытия дня по филиалам')

        self._close_day(main_win=main_win)

        self.notifiers.log.send_message(message=f'Ожидание окончания закрытия дня по филиалам')

        self._wait_for_day_procedure_end(
            main_win=main_win,
            file_name='close_day_all',
            procedure_type='close'
        )

        self.notifiers.log.send_message(message=f'День закрыт по филиалам')

        self._reset_to_00(main_win=main_win)

        self.notifiers.log.send_message(message=f'Начало закрытия дня по 00')

        self._close_day(main_win=main_win, main_branch_selected=True)

        self.notifiers.log.send_message(message=f'Ожидание закрытия дня по 00')

        self._wait_for_day_procedure_end(
            main_win=main_win,
            file_name='close_day_00',
            procedure_type='close',
            main_branch_selected=True
        )

        self.notifiers.log.send_message(message=f'День закрыт по 00')

        self._change_day(_date=self.today.next_work_date_str)

        self.notifiers.log.send_message(message=f'Операционный день изменен на {self.today.next_work_date_str}')

        self._refresh(_window=main_win)

        self._select_all_branches(_window=main_win)
        self._reset_to_00(main_win=main_win)
        main_win.click_input(button='left', coords=self.buttons.open_oper_day.coords, absolute=True)

        self.notifiers.log.send_message(message=f'Начало открытия дня по 00')

        self._open_day(main_win=main_win, main_branch_selected=True)

        self.notifiers.log.send_message(message=f'Ожидание открытия дня по 00')

        self._wait_for_day_procedure_end(
            main_win=main_win,
            file_name='open_day_00',
            procedure_type='open',
            main_branch_selected=True
        )

        self.notifiers.log.send_message(message=f'День открыт')

    def step8(self) -> None:

        self.notifiers.log.send_message(message=f'Начало работы в режиме SORDPAY')

        self._choose_mode(mode='SORDPAY')

        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit8'].wrapper_object().set_text(self.today.next_work_date_str)
        filter_win['Edit6'].wrapper_object().set_text(self.today.next_work_date_str)
        filter_win['Edit2'].wrapper_object().set_text('1')
        filter_win['Edit4'].wrapper_object().set_text('1')
        sleep(1)
        filter_win['OK'].wrapper_object().click()

        main_win = self._get_window(title='Расчетные документы филиала', timeout=600)
        self.utils.type_keys(main_win, '{F12}')

        template_win = self._get_window(title='Шаблоны платежей')
        self.utils.type_keys(template_win, '{F9}')

        filter_win2 = self._get_window(title='Фильтр')
        filter_win2['Edit2'].wrapper_object().type_keys('Утро~~', pause=.2)

        self.notifiers.log.send_message(message=f'Выбран документ шаблона "Утро"')

        order_win = self._get_window(title='Мемориальный ордер', timeout=360)

        remainder_sum = ''.join(re.findall(r'[\d.]', order_win['Edit4'].wrapper_object().window_text().strip()))
        order_win['Edit26'].type_keys(f'{remainder_sum}~', pause=.1)

        self.notifiers.log.send_message(message=f'Перенесена сумма {remainder_sum} в поле')

        sleep(1)
        order_win.click_input(button='left', coords=self.buttons.save.coords, absolute=True)

        self.notifiers.log.send_message(message=f'Документ "Утро" сохранен')

        sleep(4)

        order_win.wait(wait_for='active', timeout=360)
        order_win.close()
        sleep(2)
        main_win.type_keys('~')
        order_win = self._get_window(title='Мемориальный ордер', timeout=360)
        order_win.type_keys('{F4}')
        order_win['Edit58'].type_keys('13~')
        order_win.click_input(button='left', coords=self.buttons.save.coords, absolute=True)
        sleep(2)
        order_win.close()

        main_win.click_input(button='left', coords=self.buttons.operations.coords, absolute=True)
        main_win.type_keys('{DOWN}', pause=.1)
        main_win.close()
        self.notifiers.log.send_message(message=f'Сумма по документу "Утро" оплачена')

    def step9(self) -> None:
        main_win = self._get_window(title='Состояние операционных периодов')
        self._select_all_branches(_window=main_win)

        self.notifiers.log.send_message(message=f'Начало регламентной процедуры 1 по филиалам')

        main_win.click_input(button='left', coords=self.buttons.reg_procedure_1.coords, absolute=True)
        procedure_win = self._get_window(title='Регламентная процедура 1')

        self.notifiers.log.send_message(message=f'Ожидание завершения регламентной процедуры 1 по филиалам')

        self._fill_procedure_form(
            procedure_win=procedure_win,
            main_win=main_win,
            main_branch_selected=False,
            file_name='reg_procedure_1_all',
        )

        self._refresh(_window=main_win)
        self._select_all_branches(_window=main_win)
        self._reset_to_00(main_win=main_win)

        self.notifiers.log.send_message(message=f'Начало регламентной процедуры по 00')

        main_win.click_input(button='left', coords=self.buttons.reg_procedure_1.coords, absolute=True)
        procedure_win = self._get_window(title='Регламентная процедура 1')

        self.notifiers.log.send_message(message=f'Ожидание завершения регламентной процедуры 1 по 00')

        self._fill_procedure_form(
            procedure_win=procedure_win,
            main_win=main_win,
            main_branch_selected=True,
            file_name='reg_procedure_1_00',
        )

        self.notifiers.log.send_message(message=f'Ожидание завершения регламентной процедуры 1 по филиалам')

        self._refresh(_window=main_win)
        self._select_all_branches(_window=main_win)
        self._reset_to_00(main_win=main_win)

        main_win.click_input(button='left', coords=self.buttons.reg_procedure_4.coords, absolute=True)

        self.notifiers.log.send_message(message=f'Начало регламентной процедуры 00 по филиалам')

        procedure_win = self._get_window(title='Регламентная процедура 4')
        self._fill_procedure_form(
            procedure_win=procedure_win,
            main_win=main_win,
            main_branch_selected=True,
            file_name='reg_procedure_4_00',
        )

    def step10(self):
        self.notifiers.log.send_message('Начало выгрузки отчета')

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
        report_root = r'C:\Temp'
        report_name = 'PC05_101.xls'
        report_path = os.path.join(report_root, report_name)
        file_win['Edit2'].wrapper_object().set_text(text=report_root)
        file_win['Edit4'].wrapper_object().set_text(text=report_name)
        try:
            file_win['ComboBox'].wrapper_object().select(11)
        except (IndexError, ValueError):
            pass
        file_win['OK'].wrapper_object().click()

        params_win = self._get_window(title='Параметры отчета ')
        params_win['Edit2'].wrapper_object().set_text(text=self.robot_time.start_str)
        self.robot_time.update()
        params_win['Edit4'].wrapper_object().set_text(text=self.robot_time.end_str)
        params_win['OK'].wrapper_object().click()

        while not self.utils.is_correct_file(root=report_root, xls_file_name=report_name):
            if not os.path.exists(path=report_path):
                continue
            if os.path.getsize(filename=report_path) == 0:
                continue
        self.notifiers.log.send_message('Отчет успешно выгрузился')
        self.notifiers.log.send_message(message=report_path, is_document=True)

    def exists_950(self, date: str):
        self._choose_mode(mode='EXTRCT')

        filter_win = self._get_window(title='Фильтр')
        filter_win['Edit8'].wrapper_object().set_text(text=date)
        filter_win['OKButton'].wrapper_object().click()
        sleep(2)

        confirm_win = self.app.window(title='Подтверждение')
        if confirm_win.exists():
            confirm_win.close()
            filter_win.close()
            return False
        vypisky_win = self._get_window(title='Выписки')
        vypisky_win.close()
        return True

    def run(self) -> None:
        self._choose_mode(mode='COPPER')
        main_win = self._get_window(title='Состояние операционных периодов')
        self._get_buttons(main_win=main_win)

        self.step1()
        self.step2()
        self.step3()

        while not self.exists_950(self.today.date_str):
            self.notifiers.log.send_message(message=f'Ожидание выписки за {self.today.date_str} ...')
            sleep(360)

        self.step4()
        self.step5()
        self.step6()
        self.step7()
        self.step8()
        self.step9()
        self.step10()
