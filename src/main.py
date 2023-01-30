import os
import sys
import warnings
import dotenv
from data_structures import Credentials, Process
from robot import Robot
from bot_notification import TelegramNotifier
import datetime
import requests


def main(env: str) -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_usr, colvir_psw = os.getenv(f'COLVIR_USR_{env}'), os.getenv(f'COLVIR_PSW_{env}')
    process_name, process_path = 'COLVIR', os.getenv('COLVIR_PROCESS_PATH')

    session = requests.Session()

    args = {
        'credentials': Credentials(usr=colvir_usr, psw=colvir_psw),
        'process': Process(name=process_name, path=process_path),
        'notifier': TelegramNotifier(chat_id=os.getenv(f'CHAT_ID_{env}'), session=session),
        'today': datetime.date(2023, 1, 24),
        'session': session,
    }

    args['notifier'].send_notification('Робот начинает работу.')
    try:
        robot: Robot = Robot(**args)
        robot.run()
    except KeyboardInterrupt:
        session.close()
        sys.exit(-1073741510)


if __name__ == '__main__':
    main(env=sys.argv[1])
