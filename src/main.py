import os
import warnings

import dotenv
import requests

from bot_notification import TelegramNotifier
from data_structures import Credentials, Notifiers, Process
from robot import Robot


def main() -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_usr, colvir_psw = os.getenv(f'COLVIR_USR'), os.getenv(f'COLVIR_PSW')
    process_name, process_path = 'COLVIR', os.getenv('COLVIR_PROCESS_PATH')

    with requests.Session() as session:
        credentials = Credentials(usr=colvir_usr, psw=colvir_psw)
        process = Process(name=process_name, path=process_path)
        notifiers = Notifiers(
            log=TelegramNotifier(token=os.getenv('TOKEN_LOG'), chat_id=os.getenv(f'CHAT_ID_LOG'), session=session),
            alert=TelegramNotifier(token=os.getenv('TOKEN_ALERT'), chat_id=os.getenv(f'CHAT_ID_ALERT'), session=session)
        )

        # notifiers.log.send_message('Робот начинает работу')
        robot: Robot = Robot(
            credentials=credentials,
            process=process,
            notifiers=notifiers,
            session=session
        )
        robot.run()
        # notifiers.log.send_message('Робот успешно окончил работу')


if __name__ == '__main__':
    main()
