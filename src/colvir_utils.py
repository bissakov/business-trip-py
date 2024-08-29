import re
from time import sleep
from typing import Tuple

import pywinauto
import win32con
import win32gui
from pywinauto import mouse, win32functions

from src.process_utils import kill_all_processes


class Colvir:
    def __init__(self, process_path: str, user: str, password: str):
        self.process_path = process_path
        self.user = user
        self.password = password
        self.app = self.open_colvir()

    def open_colvir(self) -> pywinauto.Application:
        app = None
        for _ in range(10):
            try:
                app = pywinauto.Application().start(cmd_line=self.process_path)
                self.login(app=app, user=self.user, password=self.password)
                self.check_interactivity(app=app)
                break
            except pywinauto.findwindows.ElementNotFoundError:
                if self.change_password(app=app):
                    break
                kill_all_processes("COLVIR")
                continue

        assert app is not None, Exception("max_retries exceeded")
        return app

    @staticmethod
    def change_password(app: pywinauto.Application) -> None:
        attention_win = app.window(title="Внимание")
        if not attention_win.exists():
            return
        attention_win["OK"].click()

        change_password_win = app.window(title_re="Смена пароля.+")
        change_password_win["Edit0"].set_text("ROBOTIZ2024_")
        change_password_win["Edit2"].set_text("ROBOTIZ2024_")
        change_password_win["OK"].click()

        confirm_win = app.window(title="Colvir Banking System", found_index=1)
        confirm_win.send_keystrokes("{ENTER}")

        mode_win = app.window(title="Выбор режима")
        return mode_win.exists()

    @staticmethod
    def login(app: pywinauto.Application, user: str, password: str) -> None:
        if not user or not password:
            raise ValueError("COLVIR_USR or COLVIR_PSW is not set")

        login_win = app.window(title="Вход в систему")

        login_username = login_win["Edit2"]
        login_password = login_win["Edit"]

        login_username.set_text(text=user)
        if login_username.window_text() != user:
            login_username.set_text("")
            login_username.type_keys(user, set_foreground=False)

        login_password.set_text(text=password)
        if login_password.window_text() != password:
            login_password.set_text("")
            login_password.type_keys(password, set_foreground=False)

        login_win["OK"].click()

        sleep(1)
        if login_win.exists() and app.window(title="Произошла ошибка").exists():
            raise pywinauto.findwindows.ElementNotFoundError()

    @staticmethod
    def check_interactivity(app: pywinauto.Application) -> None:
        choose_mode(app=app, mode="TREPRT")
        sleep(1)

        close_window(win=app.window(title="Выбор отчета"), raise_error=True)

    def get_app(self) -> pywinauto.Application:
        assert self.app is not None
        return self.app


def set_focus_win32(win: pywinauto.WindowSpecification) -> None:
    if win.wrapper_object().has_focus():
        return

    handle = win.wrapper_object().handle

    mouse.move(coords=(-10000, 500))
    if win.is_minimized():
        if win.was_maximized():
            win.maximize()
        else:
            win.restore()
    else:
        win32gui.ShowWindow(handle, win32con.SW_SHOW)
    win32gui.SetForegroundWindow(handle)

    win32functions.WaitGuiThreadIdle(handle)


def set_focus(win: pywinauto.WindowSpecification, retries: int = 20) -> None:
    while retries > 0:
        try:
            if retries % 2 == 0:
                set_focus_win32(win)
            else:
                win.set_focus()
            break
        except (Exception, BaseException):
            retries -= 1
            sleep(5)
            continue

    if retries <= 0:
        raise Exception("Failed to set focus")


def press(win: pywinauto.WindowSpecification, key: str, pause: float = 0) -> None:
    set_focus(win)
    win.type_keys(key, pause=pause, set_foreground=False)


def choose_mode(app: pywinauto.Application, mode: str) -> None:
    mode_win = app.window(title="Выбор режима")
    mode_win["Edit2"].set_text(text=mode)
    press(mode_win["Edit2"], "~")


def close_window(win: pywinauto.WindowSpecification, raise_error: bool = False) -> None:
    if win.exists():
        win.close()
        return

    if raise_error:
        raise pywinauto.findwindows.ElementNotFoundError(f"Window {win} does not exist")


def get_window(
    app: pywinauto.Application,
    title: str,
    wait_for: str = "exists",
    timeout: int = 20,
    regex: bool = False,
    found_index: int = 0,
) -> pywinauto.WindowSpecification:
    window = (
        app.window(title=title, found_index=found_index)
        if not regex
        else app.window(title_re=title, found_index=found_index)
    )
    window.wait(wait_for=wait_for, timeout=timeout)
    sleep(0.5)
    return window


def type_keys(
    window: pywinauto.WindowSpecification,
    keystrokes: str,
    step_delay: float = 0.1,
    delay_after: float = 0.5,
) -> None:
    set_focus(window)
    for command in list(filter(None, re.split(r"({.+?})", keystrokes))):
        try:
            window.type_keys(command, set_foreground=False)
        except pywinauto.base_wrapper.ElementNotEnabled:
            sleep(1)
            window.type_keys(command, set_foreground=False)
        sleep(step_delay)

    sleep(delay_after)
