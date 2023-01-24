import os
import sys
import warnings
import dotenv
from data_structures import Credentials, Process
from robot import Robot


def main(env: str) -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_usr, colvir_psw = os.getenv(f'COLVIR_USR_{env}'), os.getenv(f'COLVIR_PSW_{env}')
    process_name, process_path = 'COLVIR', os.getenv('COLVIR_PROCESS_PATH')

    credentials: Credentials = Credentials(usr=colvir_usr, psw=colvir_psw)
    process: Process = Process(name=process_name, path=process_path)

    robot: Robot = Robot(credentials=credentials, process=process)
    robot.run()


if __name__ == '__main__':
    main(env=sys.argv[1])
