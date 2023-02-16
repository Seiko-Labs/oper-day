import os
from datetime import datetime, timedelta
from typing import List

import dotenv
import oracledb
import requests

from colvir import Colvir
from data_structures import Credentials, DateInfo, Notifiers, Process, RobotWorkTime, WorkStatus
from utils import Utils
from work_calendar import CalendarScraper


class Robot:
    def __init__(self, credentials: Credentials, process: Process, notifiers: Notifiers,
                 session: requests.Session, today: datetime.date = datetime.now().date()) -> None:
        self.restricted_pids: List[int] = []
        dotenv.load_dotenv()
        self.utils = Utils()
        self.session = session

        self.credentials = credentials
        self.process = process
        self.notifiers = notifiers
        self.today = DateInfo(date=today)
        self.robot_time = RobotWorkTime(start=datetime.now())

    def is_work_day(self) -> bool:
        scraper = CalendarScraper(
            year=self.today.date.year,
            backup_file=fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{self.today.date.year}.html',
            session=self.session
        )
        date_infos: List[DateInfo] = scraper.date_infos

        work_status: int = scraper.get_work_status(today=self.today.date, dates=date_infos)

        i = 1
        while True:
            date = self.today.date + timedelta(days=i)
            if scraper.get_work_status(today=date, dates=date_infos) == WorkStatus.WORK:
                self.today.next_work_date_str = DateInfo(date=date).date_str
                break
            i += 1

        if work_status == WorkStatus.LONG:
            self.today = DateInfo(date=self.today.date - timedelta(days=1), is_work_day=False)

        return True if work_status in (WorkStatus.WORK, WorkStatus.LONG) else False

    @staticmethod
    def emergency_call() -> None:
        oracledb.init_oracle_client()

        config_dir: str = r'C:\app\client\robot.ad\product\12.2.0\client_1\network\admin'

        with oracledb.connect(
            user=os.getenv('ORACLE_USR'), password=os.getenv('ORACLE_PSW'),
            host=os.getenv('ORACLE_HOST'), port=int(os.getenv('ORACLE_PORT')),
            service_name='GCTI8TST', config_dir=config_dir,
        ) as connection:

            cursor = connection.cursor()

            last_record_id, chain_id = cursor.execute('''
                SELECT RECORD_ID, CHAIN_ID
                FROM ocs.opovesheniye_robotom_po_colvir
                ORDER BY RECORD_ID DESC
            ''').fetchone()
            phone_number: str = '989079000000'

            cursor.execute(f'''
                INSERT INTO ocs.opovesheniye_robotom_po_colvir
                    (RECORD_ID, CONTACT_INFO, CONTACT_INFO_TYPE, RECORD_TYPE, RECORD_STATUS,
                    CALL_RESULT, ATTEMPT, DAILY_FROM, DAILY_TILL, TZ_DBID, CHAIN_ID, CHAIN_N)
                VALUES
                    ({last_record_id + 1}, {phone_number}, 1, 2, 1, 28, 0, 39600, 72000, 134, {chain_id + 1}, 0)
            ''')

            connection.commit()

    def run(self) -> None:
        if not self.is_work_day():
            self.notifiers.log.send_message(message='Не рабочий день. Завершаем работу.')
            return

        self.utils.kill_all_processes(proc_name='COLVIR', restricted_pids=self.restricted_pids)

        colvir = Colvir(
            credentials=self.credentials,
            process=self.process,
            today=self.today,
            robot_time=self.robot_time,
            notifiers=self.notifiers
        )

        try:
            colvir.run()
        except RuntimeError as e:
            self.notifiers.alert.send_message(message='Не удалось запустить Colvir')
            self.emergency_call()
            raise e
        except Exception as e:
            self.notifiers.alert.send_message(message='Случилась непредвиденная ошибка')
            self.emergency_call()
            raise e
