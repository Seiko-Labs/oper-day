import datetime
from dataclasses import dataclass


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
