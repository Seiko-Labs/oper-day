import os
from datetime import datetime
from typing import List, Dict, Any
import dotenv
import oracledb
from colvir import Colvir
from data_structures import Notifiers, Credentials, Process, DateInfo, RobotWorkTime
from utils import Utils


class Robot:
    def __init__(self, credentials: Credentials, process: Process, notifiers: Notifiers, today: datetime.date) -> None:
        self.restricted_pids: List[int] = []
        dotenv.load_dotenv()
        self.utils = Utils()

        self.credentials = credentials
        self.process = process
        self.notifiers = notifiers
        self.today = DateInfo(date=today)
        self.robot_time = RobotWorkTime(start=datetime.now())

    @staticmethod
    def emergency_call() -> None:
        oracledb.init_oracle_client()
        params: Dict[str, Any] = {
            'user': os.getenv('ORACLE_USR'),
            'password': os.getenv('ORACLE_PSW'),
            'host': os.getenv('ORACLE_HOST'),
            'port': os.getenv('ORACLE_PORT'),
            'service_name': 'GCTI8TST',
            'config_dir': r'C:\app\client\robot.ad\product\12.2.0\client_1\network\admin',
        }
        with oracledb.connect(**params) as connection:
            cursor = connection.cursor()

            last_record_id: int = cursor.execute(
                'SELECT RECORD_ID FROM ocs.opovesheniye_robotom_po_colvir ORDER BY RECORD_ID DESC').fetchone()[0]
            columns: str = 'RECORD_ID, CONTACT_INFO, CONTACT_INFO_TYPE, RECORD_TYPE, RECORD_STATUS,' \
                           'CALL_RESULT, ATTEMPT, DAILY_FROM, DAILY_TILL, TZ_DBID, CHAIN_ID, CHAIN_N'
            phone_number: str = '988079000000'

            cursor.execute(f'''
                INSERT INTO ocs.opovesheniye_robotom_po_colvir ({columns})
                VALUES ({last_record_id + 1}, {phone_number}, 1, 2, 1, 28, 0, 39600, 72000, 134, 60, 0)
            ''')

            connection.commit()

    def run(self) -> None:
        self.utils.kill_all_processes(proc_name='COLVIR', restricted_pids=self.restricted_pids)

        colvir: Colvir = Colvir(
            credentials=self.credentials,
            process=self.process,
            today=self.today,
            robot_time=self.robot_time,
            notifiers=self.notifiers
        )
        try:
            colvir.run()
        except (RuntimeError, Exception) as e:
            self.notifiers.alert.send_message(message='Не удалось запустить Colvir')
            # self.emergency_call()
            raise e
