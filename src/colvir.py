from time import sleep
from typing import List
import psutil
from pywinauto import Desktop, Application
from pywinauto.application import ProcessNotFoundError
from pywinauto.controls.hwndwrapper import DialogWrapper
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
from pywinauto.timings import TimeoutError as TimingsTimeoutError
from actions import Actions
from data_structures import Notifiers, Credentials, Process, DateInfo, RobotWorkTime
from utils import Utils


class Colvir:
    def __init__(self, credentials: Credentials, process: Process, today: DateInfo,
                 robot_time: RobotWorkTime, notifiers: Notifiers) -> None:
        self.credentials: Credentials = credentials
        self.process = process
        self.pid: int or None = None
        self.app: Application or None = None
        self.utils: Utils = Utils()
        self.today = today
        self.robot_time = robot_time
        self.notifiers = notifiers
        self.retry_count = 0

    def run(self) -> None:
        if self.retry_count == 3:
            raise RuntimeError('Не удалось запустить Colvir')

        try:
            self.notifiers.log.send_message(message='Запуск Colvir')
            Application(backend='win32').start(cmd_line=self.process.path)
            self.login()
            sleep(4)
        except (ElementNotFoundError, TimingsTimeoutError):
            self.retry()
            return
        try:
            self.pid: int = self.utils.get_current_process_pid(proc_name='COLVIR')
            self.app: Application = Application(backend='win32').connect(process=self.pid)
            try:
                if self.app.Dialog.window_text() == 'Произошла ошибка':
                    self.retry()
                    return
            except MatchError:
                pass
        except ProcessNotFoundError:
            sleep(1)
            self.pid: int = self.utils.get_current_process_pid(proc_name='COLVIR')
            self.app: Application = Application(backend='win32').connect(process=self.pid)
        try:
            self.notifiers.log.send_message(message='Успешный вход в Colvir')
            self.confirm_warning()
            sleep(1)
        except (ElementNotFoundError, MatchError):
            self.retry()
            return
        actions = Actions(
            app=self.app,
            today=self.today,
            robot_time=self.robot_time,
            notifiers=self.notifiers,
        )
        actions.run()

    def login(self) -> None:
        desktop: Desktop = Desktop(backend='win32')
        try:
            login_win = desktop.window(title='Вход в систему')
            login_win.wait(wait_for='exists', timeout=20)
            login_win['Edit2'].wrapper_object().set_text(text=self.credentials.usr)
            login_win['Edit'].wrapper_object().set_text(text=self.credentials.psw)
            login_win['OK'].wrapper_object().click()
        except ElementAmbiguousError:
            windows: List[DialogWrapper] = Desktop(backend='win32').windows()
            for win in windows:
                if 'Вход в систему' not in win.window_text():
                    continue
                self.utils.kill_process(pid=win.process_id())
            raise ElementNotFoundError

    def confirm_warning(self) -> None:
        found = False
        for window in self.app.windows():
            if found:
                break
            if window.window_text() != 'Colvir Banking System':
                continue
            win = self.app.window(handle=window.handle)
            for child in win.descendants():
                if child.window_text() == 'OK':
                    found = True
                    child.send_keystrokes('{ENTER}')
                    break

    def kill(self) -> None:
        try:
            self.utils.kill_process(pid=self.pid)
        except psutil.NoSuchProcess:
            self.pid: int = self.utils.get_current_process_pid(proc_name='COLVIR')
            self.utils.kill_process(pid=self.pid)

    def retry(self) -> None:
        self.retry_count += 1
        self.kill()
        self.run()
