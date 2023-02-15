import os
import warnings
from typing import List
import dotenv
import requests
from datetime import datetime, timedelta
from data_structures import Credentials, Process, Notifiers, DateInfo, WorkStatus
from robot import Robot
from bot_notification import TelegramNotifier
from work_calendar import CalendarScraper


def get_work_status(today: DateInfo, session: requests.Session) -> int:
    scraper = CalendarScraper(
        year=today.date.year,
        backup_file=fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{today.date.year}.html',
        session=session
    )
    date_infos: List[DateInfo] = scraper.date_infos

    work_status = scraper.get_work_status(today=today.date, dates=date_infos)

    i = 1
    while True:
        date = today.date + timedelta(days=i)
        if scraper.get_work_status(today=date, dates=date_infos) == WorkStatus.WORK:
            today.next_work_date_str = DateInfo(date=date).date_str
            break
        i += 1

    return work_status


def main() -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_usr, colvir_psw = os.getenv(f'COLVIR_USR'), os.getenv(f'COLVIR_PSW')
    process_name, process_path = 'COLVIR', os.getenv('COLVIR_PROCESS_PATH')

    with requests.Session() as session:
        credentials = Credentials(usr=colvir_usr, psw=colvir_psw)
        process = Process(name=process_name, path=process_path)
        today = DateInfo(datetime.now().date())
        notifiers = Notifiers(
            log=TelegramNotifier(token=os.getenv('TOKEN_LOG'), chat_id=os.getenv(f'CHAT_ID_LOG'), session=session),
            alert=TelegramNotifier(token=os.getenv('TOKEN_ALERT'), chat_id=os.getenv(f'CHAT_ID_ALERT'), session=session)
        )

        work_status = get_work_status(today=today, session=session)
        if work_status == WorkStatus.LONG:
            today = DateInfo(date=today.date - timedelta(days=1), is_work_day=False)
        if work_status == WorkStatus.HOLIDAY:
            notifiers.log.send_message(message='Не рабочий день. Завершаем работу.')
            # return

        # notifiers.log.send_message('Робот начинает работу.')
        robot: Robot = Robot(
            credentials=credentials,
            process=process,
            notifiers=notifiers,
            today=today
        )
        robot.run()


if __name__ == '__main__':
    main()
