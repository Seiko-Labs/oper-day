import os
import sys
import warnings
import dotenv
import requests
from data_structures import Credentials, Process, Notifiers
from robot import Robot
from bot_notification import TelegramNotifier


def main(env: str) -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_usr, colvir_psw = os.getenv(f'COLVIR_USR_{env}'), os.getenv(f'COLVIR_PSW_{env}')
    process_name, process_path = 'COLVIR', os.getenv('COLVIR_PROCESS_PATH')

    session = requests.Session()

    credentials = Credentials(usr=colvir_usr, psw=colvir_psw)
    process = Process(name=process_name, path=process_path)
    notifiers = Notifiers(
        log=TelegramNotifier(token=os.getenv('TOKEN_LOG'), chat_id=os.getenv(f'CHAT_ID_LOG'), session=session),
        alert=TelegramNotifier(token=os.getenv('TOKEN_ALERT'), chat_id=os.getenv(f'CHAT_ID_ALERT'), session=session)
    )
    # today = datetime.date(2023, 2, 10)
    session = session

    notifiers.log.send_message('Робот начинает работу.')
    try:
        robot: Robot = Robot(
            credentials=credentials,
            process=process,
            notifiers=notifiers,
            session=session,
        )
        robot.run()
    except KeyboardInterrupt:
        session.close()
        sys.exit(-1073741510)


if __name__ == '__main__':
    main(env=sys.argv[1])
