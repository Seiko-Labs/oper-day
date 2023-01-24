from typing import List
import dotenv
from bot_notification import TelegramNotifier
from colvir import Colvir
from data_structures import Credentials, Process, DateInfo, WorkStatus
from utils import Utils
from work_calendar import CalendarScraper
from datetime import datetime, timedelta
from datetime import date


class Robot:
    def __init__(self, credentials: Credentials, process: Process, _date: datetime.date = datetime.now().date()) -> None:
        self.credentials: Credentials = credentials
        self.process: Process = process
        self.restricted_pids: List[int] = []
        dotenv.load_dotenv()
        self.notifier = TelegramNotifier()
        self.utils = Utils()
        self.today = DateInfo(date=_date)

    def is_work_day(self) -> bool:
        scraper = CalendarScraper(
            year=self.today.date.year,
            backup_file=fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{self.today.date.year}.html'
        )
        date_infos: List[DateInfo] = scraper.date_infos

        work_status = scraper.get_work_status(today=self.today.date, dates=date_infos)

        if work_status == WorkStatus.LONG:
            self.today = DateInfo(date=self.today.date - timedelta(days=1), is_work_day=False)

        return True if work_status in (WorkStatus.WORK, WorkStatus.LONG) else False

    def run(self) -> None:
        if not self.is_work_day():
            return

        self.utils.kill_all_processes(proc_name='COLVIR', restricted_pids=self.restricted_pids)
        colvir: Colvir = Colvir(credentials=self.credentials, process=self.process, today=self.today)
        colvir.run()
