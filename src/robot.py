from typing import List
import dotenv
from bot_notification import TelegramNotifier
from colvir import Colvir
from data_structures import Credentials, Process, DateInfo, WorkStatus, RobotWorkTime
from utils import Utils
from work_calendar import CalendarScraper
from datetime import datetime, timedelta
from datetime import date


class Robot:
    def __init__(self, credentials: Credentials, process: Process, _date: datetime.date = datetime.now().date()) -> None:
        self.restricted_pids: List[int] = []
        dotenv.load_dotenv()
        self.notifier = TelegramNotifier()
        self.utils = Utils()
        self.args = {
            'credentials': credentials,
            'process': process,
            'today': DateInfo(date=_date),
            'robot_time': RobotWorkTime(start=datetime.now())
        }

    def is_work_day(self) -> bool:
        today = self.args['today']
        scraper = CalendarScraper(
            year=today.date.year,
            backup_file=fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{today.date.year}.html'
        )
        date_infos: List[DateInfo] = scraper.date_infos

        work_status = scraper.get_work_status(today=today.date, dates=date_infos)

        if work_status == WorkStatus.LONG:
            self.args['today'] = DateInfo(date=today.date - timedelta(days=1), is_work_day=False)

        return True if work_status in (WorkStatus.WORK, WorkStatus.LONG) else False

    def run(self) -> None:
        if not self.is_work_day():
            return

        self.utils.kill_all_processes(proc_name='COLVIR', restricted_pids=self.restricted_pids)
        colvir: Colvir = Colvir(**self.args)
        colvir.run()
