import warnings
from data_structures import Credentials, Process
from robot import Robot


def main(env: str) -> None:
    warnings.simplefilter(action='ignore', category=UserWarning)

    if env == 'dev':
        colvir_usr, colvir_psw = 'colvir', 'colvir147'
        process_name, process_path = 'COLVIR', r'C:\CBS_T_new\COLVIR.EXE'
    else:
        colvir_usr, colvir_psw = 'robot', 'Asd_24-08-2022'
        process_name, process_path = 'COLVIR', r'C:\CBS_R\COLVIR.EXE'

    credentials: Credentials = Credentials(usr=colvir_usr, psw=colvir_psw)
    process: Process = Process(name=process_name, path=process_path)

    robot: Robot = Robot(credentials=credentials, process=process)
    robot.run()


if __name__ == '__main__':
    main(env='dev')
