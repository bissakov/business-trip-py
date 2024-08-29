import json
import os
import warnings
from compileall import compile_file
from datetime import datetime, timedelta
from time import sleep
from typing import List, cast, Optional

import dotenv
import pandas as pd
import pywinauto.base_wrapper
from pywinauto import mouse, WindowSpecification
from pywinauto.win32structures import RECT

from src.colvir_utils import get_window, Colvir, choose_mode
from src.data import Buttons, Button, Order
from src.excel_utils import xls_to_xlsx
from src.process_utils import kill_all_processes
from src.wiggle import wiggle_mouse
import pywinauto.timings


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in .env"
    return value


def load_orders(order_json_path: str) -> List[Order]:
    with open(order_json_path, "r", encoding="utf-8") as f:
        orders_json = json.load(f)

    orders: List[Order] = []
    for order_json in orders_json:
        order = Order(
            employee_fullname=order_json["employee_fullname"],
            employee_names=order_json["employee_names"],
            order_number=order_json["order_number"],
            sign_date=order_json["sign_date"],
            start_date=order_json["start_date"],
            end_date=order_json["end_date"],
            trip_place=order_json["trip_place"],
            trip_target=order_json["trip_target"],
            main_order_number=order_json["main_order_number"],
            main_order_start_date=order_json["main_order_start_date"],
            deputy_fullname=order_json["deputy_fullname"],
            deputy_names=order_json["deputy_names"],
        )
        orders.append(order)

    return orders


def create_report(report_file_path: str):
    if not os.path.exists(report_file_path):
        df = pd.DataFrame(
            {
                "Дата": [],
                "Сотрудник": [],
                "Операция": [],
                "Номер приказа": [],
                "Статус": [],
            }
        )
        df.to_excel(report_file_path, index=False)


def update_report(
    person_name: str,
    order: Order,
    report_file_path: str,
    today: str,
    operation: str,
    status: str,
):
    order_number = order.order_number

    df = pd.read_excel(report_file_path)

    if not (
        (df["Дата"] == today)
        & (df["Сотрудник"] == person_name)
        & (df["Операция"] == operation)
        & (df["Номер приказа"] == order_number)
    ).any():
        new_row = {
            "Дата": today,
            "Сотрудник": person_name,
            "Операция": operation,
            "Номер приказа": order_number,
            "Статус": status,
        }
        df.loc[len(df)] = new_row
        df.to_excel(report_file_path, index=False)


def close_dialog(app: pywinauto.Application) -> None:
    dialog_win = get_window(app=app, title="Colvir Banking System", found_index=0)
    dialog_win.set_focus()
    sleep(0.5)
    dialog_win["OK"].click_input()


def change_oper_day(app: pywinauto.Application, start_date: str):
    start_date = datetime.strptime(start_date, "%d%m%Y").strftime("%d.%m.%y")
    current_oper_day_win = get_window(app=app, title="Текущий операционный день")
    current_oper_day_win["Edit2"].set_text(start_date)
    current_oper_day_win["OK"].click()
    attention_win = app.window(title="Внимание")
    if not attention_win.exists():
        start_date_dt = datetime.strptime(start_date, "%d.%m.%y")
        start_date = (start_date_dt - timedelta(days=1)).strftime("%d%m%Y")
        close_dialog(app=app)
        change_oper_day(app=app, start_date=start_date)
        return

    attention_win["&Да"].click()
    current_oper_day_win["OK"].click()
    sleep(0.5)
    close_dialog(app=app)


def save_excel(app: pywinauto.Application, work_folder: str):
    file_win = get_window(app=app, title="Выберите файл для экспорта")

    orders_file_path = os.path.join(work_folder, "orders.xls")
    orders_xlsx_file_path = os.path.join(work_folder, "orders.xlsx")

    file_win["Edit4"].set_text(orders_file_path)
    file_win["&Save"].click_input()

    sleep(1)
    confirm_win = app.window(title="Confirm Save As")
    if confirm_win.exists():
        confirm_win["Yes"].click()

    sort_win = get_window(app=app, title="Сортировка")
    sort_win["OK"].click()

    while not os.path.exists(orders_file_path):
        sleep(5)
    sleep(1)

    kill_all_processes("EXCEL")

    xls_to_xlsx(orders_file_path, orders_xlsx_file_path)

    return orders_xlsx_file_path


def get_colvir_city_code(trip_place: str, work_folder: str) -> Optional[str]:
    with open(os.path.join(work_folder, "cities.json"), "r", encoding="utf-8") as f:
        cities = json.load(f)

    city_bpm = trip_place.replace("город ", "").replace("г. ", "")
    city_bpm = city_bpm.split(",")[0]
    city_colvir = cities.get(city_bpm)
    if not city_colvir:
        return None
    city_colvir = city_colvir.replace(f".{city_bpm}", "")

    return city_colvir


def persistent_win_exists(
    app: pywinauto.Application, title_re: str, timeout: float
) -> bool:
    try:
        app.window(title_re=title_re).wait(wait_for="enabled", timeout=timeout)
    except pywinauto.timings.TimeoutError:
        return False
    return True


def get_city_mappings(
    app: pywinauto.Application, order: Order, buttons: Buttons
) -> None:
    choose_mode(app=app, mode="PRS")
    filter_win = get_window(app=app, title="Фильтр")
    buttons.clear_form.find_and_click_button(
        app=app,
        window=filter_win,
        toolbar=filter_win["Static3"],
        target_button_name="Очистить фильтр",
    )

    filter_win["Edit8"].set_text("001")
    filter_win["Edit4"].set_text(order.employee_names[0])
    filter_win["Edit2"].set_text(order.employee_names[1])
    filter_win["OK"].click()

    personal_win = get_window(app=app, title="Персонал")
    buttons.employee_orders.find_and_click_button(
        app=app,
        window=personal_win,
        toolbar=personal_win["Static4"],
        target_button_name="Приказы по сотруднику",
    )

    orders_win = get_window(app=app, title="Приказы сотрудника")
    buttons.create_new_order.find_and_click_button(
        app=app,
        window=orders_win,
        toolbar=orders_win["Static4"],
        target_button_name="Создать новую запись (Ins)",
    )

    order_win = get_window(app=app, title="Приказ")

    order_win["Edit18"].type_keys("ORD_TRP", pause=0.1)
    order_win["Edit18"].type_keys("{TAB}")
    sleep(0.5)
    if (error_win := app.window(title="Произошла ошибка")).exists():
        error_win.close()
        order_win["Edit38"].type_keys("{TAB}")
    sleep(1)

    order_win["Edit40"].type_keys(order.order_number, pause=0.1)

    sleep(1)

    order_win["Edit4"].click_input()
    order_win["Edit4"].type_keys("001", pause=0.2)
    order_win.type_keys("{TAB}", pause=1)
    order_win["Edit10"].click_input()
    order_win["Edit10"].type_keys("0975", pause=0.2)
    order_win.type_keys("{TAB}", pause=1)

    if buttons.cities_menu.x == -1 or buttons.cities_menu.y == -1:
        order_win.set_focus()
        rect: RECT = order_win["Edit28"].rectangle()
        mid_point = rect.mid_point()

        start_point = rect.right
        end_point = rect.right + 200

        x, y = rect.right, mid_point.y
        mouse.move(coords=(x, y))

        x_offset = 5

        i = 0
        while (
            not persistent_win_exists(
                app=app, title_re="Страны и города.+", timeout=0.1
            )
            or x >= end_point
        ):
            x = start_point + i * 5
            mouse.click(button="left", coords=(x, y))
            i += 1

        buttons.cities_menu.x = x + x_offset
        buttons.cities_menu.y = y

        cities_win = app.window(title="Страны и города (командировки)")
        cities_win.close()

    mappings = {}

    order_win.set_focus()
    buttons.cities_menu.click()
    i = 0
    while i < 500:
        cities_win = get_window(app=app, title="Страны и города (командировки)")

        if i > 0:
            cities_win.type_keys("{DOWN}")
        ok_button = cities_win["OK"]
        if ok_button.is_enabled():
            ok_button.click()
        else:
            cities_win.type_keys("{ENTER}")
            continue

        cities_win.close()

        fullname = order_win["Edit18"].window_text().strip()
        colvir_name = order_win["Edit28"].window_text().strip()
        mappings[fullname] = colvir_name

        print(
            order_win["Edit28"].window_text().strip(),
            "&&",
            order_win["Edit18"].window_text().strip(),
        )

        i += 1

        buttons.cities_menu.click()

    pass


def main():
    # work_folder = {{PATH}}
    # today = {{TODAY}}
    # colvir_user = {{cred_colvir}}['username']
    # colvir_password = {{cred_colvir}}['password']

    # NOTE: Конфигурация
    project_folder = os.path.dirname(os.path.dirname(__file__))
    dotenv.load_dotenv(os.path.join(project_folder, ".env.test"))

    colvir_path = get_from_env("COLVIR_PATH")
    colvir_user = get_from_env("COLVIR_USER")
    colvir_password = get_from_env("COLVIR_PASSWORD")

    today = datetime.now().strftime("%d.%m.%y")

    # work_folder = r"C:\Users\robotX4\Desktop\entering_data_on _orders\Командировки"
    work_folder = os.path.join(project_folder, "data", "reports")

    report_file_path = os.path.join(work_folder, f"Отчет_командировки_{today}.xlsx")
    # NOTE: Создание пустого отчета по работе робота
    create_report(report_file_path)

    # Не отображать предупреждения pywinauto о 32-битном приложении
    warnings.simplefilter(action="ignore", category=UserWarning)
    kill_all_processes(proc_name="COLVIR")

    # Координаты кнопок с динамичными местоположениями
    buttons = Buttons()

    colvir = Colvir(
        process_path=colvir_path, user=colvir_user, password=colvir_password
    )
    app = colvir.app

    # Сбор приказов ранее выгружженых из BPM
    orders_json_path = os.path.join(work_folder, f"orders_{today}.json")
    orders: List[Order] = load_orders(orders_json_path)

    # get_city_mappings(app=app, order=orders[0], buttons=buttons)  #  Сбор маппингов (не запускать просто так)

    for i, order in enumerate(orders):
        # Смена операционного дня
        choose_mode(app=app, mode="TOPERD")
        change_oper_day(app=app, start_date=order.start_date)

        # Переход в Персонал (PRS)
        choose_mode(app=app, mode="PRS")
        filter_win = get_window(app=app, title="Фильтр")
        buttons.clear_form.find_and_click_button(
            app=app,
            window=filter_win,
            toolbar=filter_win["Static3"],
            target_button_name="Очистить фильтр",
        )

        # Фильтр по имени и фамилии сотрудника
        filter_win["Edit8"].set_text("001")
        filter_win["Edit4"].set_text(order.employee_names[0])
        filter_win["Edit2"].set_text(order.employee_names[1])
        filter_win["OK"].click()

        sleep(1)
        # Данное окно выходит только в случае ненахождения сотрудника
        # Записываем в отчет и идем дальше
        confirm_win = app.window(title="Подтверждение")
        if confirm_win.exists():
            update_report(
                person_name=order.employee_fullname,
                order=order,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status="Приказ не найден",
            )
            confirm_win.close()
            filter_win.close()
            personal_win = app.window(title="Персонал")
            if personal_win.exists():
                personal_win.close()
            continue

        personal_win = get_window(app=app, title="Персонал")
        # Переход в список приказов
        buttons.employee_orders.find_and_click_button(
            app=app,
            window=personal_win,
            toolbar=personal_win["Static4"],
            target_button_name="Приказы по сотруднику",
        )

        orders_win = get_window(app=app, title="Приказы сотрудника")
        orders_win.menu_select("#4->#4->#1")
        orders_file_path = save_excel(app=app, work_folder=work_folder)

        df = pd.read_excel(orders_file_path, skiprows=1)

        order_exists = (
            (df["Вид приказа"] == "Приказ о отправке работника в командировку")
            & (df["Номер приказа"] == order.order_number)
        ).any()

        # Если приказ уже существуем, идем дальше
        if order_exists:
            orders_win.close()
            personal_win.close()

            update_report(
                person_name=order.employee_fullname,
                order=order,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status="Приказ уже создан",
            )
            continue

        personal_win.set_focus()
        sleep(1)
        personal_win.type_keys("{ENTER}")

        # Переход в карточку сотрудника. Уволенных, командировачных и отпускных пропускаем
        employee_card = get_window(app=app, title="Карточка сотрудника")
        employee_status = employee_card["Edit30"].window_text().strip()
        print(order.employee_fullname, employee_status)
        if (
            employee_status == "Уволен"
            or employee_status == "В командировке"
            or employee_status == "В отпуске"
        ):
            employee_card.close()
            orders_win.close()
            personal_win.close()
            update_report(
                person_name=order.employee_fullname,
                order=order,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status=f"Невозможно создать приказ для сотрудника "
                f'со статусом "{employee_status}"',
            )
            continue

        # Сохранение подразделения и табельного номера сотрудника
        branch_num = employee_card["Edit60"].window_text()
        tab_num = employee_card["Edit34"].window_text()
        employee_card.close()

        orders_win.set_focus()
        sleep(1)
        buttons.create_new_order.find_and_click_button(
            app=app,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Создать новую запись (Ins)",
        )

        order_win = get_window(app=app, title="Приказ")

        order_win["Edit18"].type_keys("ORD_TRP", pause=0.1)
        order_win["Edit18"].type_keys("{TAB}")
        sleep(0.5)
        if (error_win := app.window(title="Произошла ошибка")).exists():
            error_win.close()
            order_win["Edit38"].type_keys("{TAB}")

        sleep(1)

        order_win["Edit40"].type_keys(order.order_number, pause=0.1)

        sleep(1)

        order_win["Edit4"].click_input()
        order_win["Edit4"].type_keys(branch_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)
        order_win["Edit10"].click_input()
        order_win["Edit10"].type_keys(tab_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)

        if not order_win.wrapper_object().has_focus():
            order_win.set_focus()
        start_date = datetime.strptime(order.start_date, "%d%m%Y").strftime("%d.%m.%y")
        end_date = datetime.strptime(order.end_date, "%d%m%Y").strftime("%d.%m.%y")

        order_win["Edit22"].click_input()
        order_win["Edit22"].set_text(start_date)

        order_win["Edit24"].click_input()
        order_win["Edit24"].set_text(end_date)

        city_code = get_colvir_city_code(
            trip_place=order.trip_place, work_folder=work_folder
        )

        if city_code is None:
            update_report(
                person_name=order.employee_fullname,
                order=order,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status=f"Не удалось заполнить приказ. Требуется проверка специалистом. "
                f"Неизвестный город/местоположение - {order.trip_place}",
            )
            order_win.type_keys("{ESC}")

            confirm_win = get_window(app=app, title="Подтверждение")
            confirm_win["&Нет"].click()

            orders_win.close()
            personal_win.close()
            continue

        order_win["Edit28"].type_keys(city_code, pause=0.2)
        order_win["Edit28"].click_input()
        order_win.type_keys("{TAB}", pause=1)

        order_win["Edit16"].type_keys(order.trip_target, pause=0.1, with_spaces=True)
        order_win["Edit16"].click_input()
        order_win.type_keys("{TAB}", pause=1)

        buttons.order_save.find_and_click_button(
            app=app,
            window=order_win,
            toolbar=order_win["Static3"],
            target_button_name="Сохранить изменения (PgDn)",
        )

        orders_win.wait(wait_for="active enabled")

        buttons.operations_list.find_and_click_button(
            app=app,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Выполнить операцию",
        )

        sleep(0.5)
        buttons.operation = Button(
            buttons.operations_list.x,
            buttons.operations_list.y + 30,
        )
        buttons.operation.check_and_click(app=app, target_button_name="Регистрация")

        registration_win = get_window(app=app, title="Подтверждение")
        registration_win["&Да"].click()
        sleep(2)
        confirm_win = app.window(title="Подтверждение")
        if confirm_win.exists():
            confirm_win.close()
        sleep(1)
        dossier_win = app.window(title="Досье сотрудника")
        if dossier_win.exists():
            dossier_win.close()

        wiggle_mouse(duration=3)

        buttons.operations_list.click()
        sleep(1)
        buttons.operation.click()
        confirm_win = get_window(app=app, title="Подтверждение")
        confirm_win["&Да"].click()
        wiggle_mouse(duration=3)

        buttons.operations_list.click()
        sleep(1)
        buttons.operation.click()
        confirm_win = get_window(app=app, title="Подтверждение")
        confirm_win["&Да"].click()
        wiggle_mouse(duration=3)

        command_win = app.window(title="Распоряжение на командировку")
        if command_win.exists():
            command_win.close()

        error_win = app.window(title="Произошла ошибка")
        if error_win.exists():
            error_msg = error_win.child_window(class_name="Edit").window_text()
            update_report(
                person_name=order.employee_fullname,
                order=order,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status=f"Не удалось ИСПОЛНИТЬ приказ. Требуется проверка специалистом. "
                f'Текст ошибки - "{error_msg}"',
            )
            error_win.close()
            orders_win.close()
            personal_win.close()
            continue

        if order.deputy_fullname is None:
            update_report(
                person_name=order.employee_fullname,
                order=order,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status="Приказ создан",
            )
            orders_win.close()
            personal_win.close()
            continue

        pass

        update_report(
            person_name=order.deputy_fullname,
            order=order,
            report_file_path=report_file_path,
            today=today,
            operation="Создание приказа",
            status=f"Приказ создан. Доплата за на период командировки сотрудника {order.employee_fullname}",
        )
        orders_win.close()
        personal_win.close()

    pass


if __name__ == "__main__":
    main()
