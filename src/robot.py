from typing import List
import dotenv
from bot_notification import TelegramNotifier
from colvir import Colvir
from data_structures import Credentials, Process
from utils import Utils


class Robot:
    def __init__(self, credentials: Credentials, process: Process) -> None:
        self.credentials: Credentials = credentials
        self.process: Process = process
        self.restricted_pids: List[int] = []
        dotenv.load_dotenv()
        self.notifier = TelegramNotifier()
        self.utils = Utils()

    def run(self) -> None:
        self.utils.kill_all_processes(proc_name='COLVIR', restricted_pids=self.restricted_pids)
        colvir: Colvir = Colvir(credentials=self.credentials, process=self.process)
        colvir.run()
