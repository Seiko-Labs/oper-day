import os
import sys
import warnings
import dotenv
from data_structures import Credentials, Process
from robot import Robot


def main(env: str) -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    process_name = 'COLVIR'
    if env == 'dev':
        colvir_usr, colvir_psw = os.getenv('COLVIR_USR_DEV'), os.getenv('COLVIR_PSW_DEV')
        process_path = r'C:\CBS_T_new\COLVIR.EXE'
    else:
        colvir_usr, colvir_psw = os.getenv('COLVIR_USR_PROD'), os.getenv('COLVIR_PSW_PROD')
        process_path = r'C:\CBS_R\COLVIR.EXE'

    credentials: Credentials = Credentials(usr=colvir_usr, psw=colvir_psw)
    process: Process = Process(name=process_name, path=process_path)

    robot: Robot = Robot(credentials=credentials, process=process)
    robot.run()


if __name__ == '__main__':
    main(env=sys.argv[1])
