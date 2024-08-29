import dataclasses
from time import sleep
from typing import Tuple, Optional

import pywinauto
from pywinauto import mouse


@dataclasses.dataclass
class Button:
    x: int = -1
    y: int = -1

    def click(self) -> None:
        mouse.click(button="left", coords=(self.x, self.y))

    def check_and_click(
        self, app: pywinauto.Application, target_button_name: str
    ) -> None:
        mouse.move(coords=(self.x, self.y))
        status_bar = app.window(title_re="Банковская система.+")["StatusBar"]
        if status_bar.window_text().strip() == target_button_name:
            self.click()

    def find_and_click_button(
        self,
        app: pywinauto.Application,
        window: pywinauto.WindowSpecification,
        toolbar: pywinauto.WindowSpecification,
        target_button_name: str,
        horizontal: bool = True,
        offset: int = 5,
    ) -> "Button":
        if self.x != -1 and self.y != -1:
            self.click()
            return self

        status_win = app.window(title_re="Банковская система.+")
        rectangle = toolbar.rectangle()
        mid_point = rectangle.mid_point()
        mouse.move(coords=(mid_point.x, mid_point.y))

        start_point = rectangle.left if horizontal else rectangle.top
        end_point = rectangle.right if horizontal else rectangle.bottom

        x, y = mid_point.x, mid_point.y
        point = 0

        x_offset = offset if horizontal else 0
        y_offset = offset if not horizontal else 0

        i = 0
        while (
            status_win["StatusBar"].window_text().strip() != target_button_name
            or point >= end_point
        ):
            point = start_point + i * 5

            if horizontal:
                x = point
            else:
                y = point

            mouse.move(coords=(x, y))
            i += 1

        window.set_focus()
        sleep(1)

        self.x = x + x_offset
        self.y = y + y_offset
        self.click()

        return self


class Buttons:
    def __init__(self):
        self.clear_form: Button = Button()
        self.employee_orders: Button = Button()
        self.create_new_order: Button = Button()
        self.order_save: Button = Button()
        self.operations_list: Button = Button()
        self.operation: Button = Button()
        self.cities_menu: Button = Button()


@dataclasses.dataclass
class Order:
    employee_fullname: str
    employee_names: Tuple[str, str]
    order_number: str
    sign_date: str
    start_date: str
    end_date: str
    trip_place: str
    trip_target: str
    main_order_number: str
    main_order_start_date: str
    deputy_fullname: Optional[str]
    deputy_names: Optional[Tuple[str, str]]
