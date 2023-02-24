import datetime
from dataclasses import dataclass
from typing import List, Tuple

from bot_notification import TelegramNotifier


@dataclass
class Credentials:
    usr: str
    psw: str


@dataclass
class Process:
    name: str
    path: str


@dataclass
class DateInfo:
    date: datetime.date or str
    is_work_day: bool = None
    date_str: str = None
    weekday: int = None
    weekday_str: str = None
    next_date_str: str = None
    prev_date_str: str = None
    next_work_date_str: str = None

    def __post_init__(self):
        if isinstance(self.date, str):
            self.date = datetime.date.fromisoformat(self.date)
        self.date_str = self.date.strftime('%d.%m.%y')
        self.weekday = self.date.weekday() + 1
        self.weekday_str = self.date.strftime('%A')
        self.next_date_str = (self.date + datetime.timedelta(days=1)).strftime('%d.%m.%y')


class WorkStatus:
    HOLIDAY = 0
    WORK = 1
    LONG = 2


@dataclass
class RobotWorkTime:
    start: datetime.datetime
    end: datetime.datetime = None
    start_str: str = None
    end_str: str = None

    def __post_init__(self):
        if not self.start_str:
            self.start_str = self.start.strftime('%d.%m.%y %H:%M')

    def update(self):
        self.end = datetime.datetime.now()
        self.end_str = self.end.strftime('%d.%m.%y %H:%M')


@dataclass
class Button:
    coords: Tuple[int, int]
    name: str
    filled: bool = False


@dataclass
class Buttons:
    open_oper_day: Button = Button(coords=(0, 0), name='Открыть новый операционный период')
    close_oper_day: Button = Button(coords=(0, 0), name='Закрыть операционный период')
    reg_procedure_1: Button = Button(coords=(0, 0), name='Регламентная процедура 1')
    reg_procedure_2: Button = Button(coords=(0, 0), name='Регламентная процедура 2')
    reg_procedure_4: Button = Button(coords=(0, 0), name='Регламентная процедура 4')
    remove_reg_procedure_4: Button = Button(coords=(0, 0), name='Снять признак выполнения регламентной процедуры 4')
    refresh: Button = Button(coords=(0, 0), name='Обновить список')
    tasks: Button = Button(coords=(0, 0), name='Все задания на обработку')
    tasks_refresh: Button = Button(coords=(0, 0), name='Обновить список')
    operations: Button = Button(coords=(0, 0), name='Выполнить операцию')
    save: Button = Button(coords=(0, 0), name='Сохранить изменения (PgDn)')
    filled_count: int = 0


@dataclass
class Notifiers:
    log: TelegramNotifier
    alert: TelegramNotifier


class EmailInfo:
    # email_list: List[str] = ['ualihanova.k@otbasybank.kz', 'sarybayeva.a@otbasybank.kz',
    #                          'Seisenbin.E@otbasybank.kz', 'musabekov.r@otbasybank.kz',
    #                          'absattarova.r@otbasybank.kz', 'mazhit.e@hcsbk.kz',
    #                          'shakhabayev.n@otbasybank.kz', 'maulenova.s@otbasybank.kz',
    #                          'ops2@otbasybank.kz']
    email_list: List[str] = ['robot.ad@hcsbk.kz']
    recepient: str = None
    subject: str = 'Протокол работы супервизоров PC05_101 TEST'
    body: str = 'Протокол работы супервизоров PC05_101 TEST'
    attachment: str = r'C:\Temp\PC05_101.xls'

    def __init__(self):
        self.recepient = ';'.join(self.email_list)
