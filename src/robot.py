from typing import List
import dotenv
import requests
from bot_notification import TelegramNotifier
from colvir import Colvir
from data_structures import Credentials, Process, DateInfo, WorkStatus, RobotWorkTime
from utils import Utils
from work_calendar import CalendarScraper
from datetime import datetime, timedelta


class Robot:
    def __init__(self, credentials: Credentials, process: Process, notifier: TelegramNotifier,
                 session: requests.Session, today: datetime.date = datetime.now().date()) -> None:
        self.restricted_pids: List[int] = []
        dotenv.load_dotenv()
        self.utils = Utils()
        self.session = session
        self.args = {
            'credentials': credentials,
            'process': process,
            'notifier': notifier,
            'today': DateInfo(date=today),
            'robot_time': RobotWorkTime(start=datetime.now())
        }

    def is_work_day(self) -> bool:
        today = self.args['today']
        scraper = CalendarScraper(
            year=today.date.year,
            backup_file=fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{today.date.year}.html',
            session=self.session
        )
        date_infos: List[DateInfo] = scraper.date_infos

        work_status = scraper.get_work_status(today=today.date, dates=date_infos)

        i = 1
        while True:
            date = today.date + timedelta(days=i)
            if scraper.get_work_status(today=date, dates=date_infos) == WorkStatus.WORK:
                self.args['today'].next_work_date_str = DateInfo(date=date).date_str
                break
            i += 1

        if work_status == WorkStatus.LONG:
            self.args['today'] = DateInfo(date=today.date - timedelta(days=1), is_work_day=False)

        return True if work_status in (WorkStatus.WORK, WorkStatus.LONG) else False

    def run(self) -> None:
        if not self.is_work_day():
            pass
            # self.args['notifier'].send_message(message='Не рабочий день. Завершаем работу.')
            # return

        self.utils.kill_all_processes(proc_name='COLVIR', restricted_pids=self.restricted_pids)
        colvir: Colvir = Colvir(**self.args)
        colvir.run()
        self.session.close()
