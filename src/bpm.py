import json
import os
from datetime import datetime

import pandas as pd
import selenium.webdriver.chrome.service as chrome_service
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait


def driver_init(executable_path: str, download_folder: str) -> Chrome:
    service = chrome_service.Service(executable_path=executable_path)
    options = ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications": 2}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": download_folder,
            "download.directory_upgrade": True,
            "download.prompt_for_download": False,
        },
    )
    driver = Chrome(service=service, options=options)
    return driver


def login(
    driver: Chrome, wait: WebDriverWait, bpm_user: str, bpm_password
) -> None:
    driver.get("https://bpmtest.kdb.kz/")

    user_input = wait.until(
        ec.presence_of_element_located((By.NAME, "u_login"))
    )
    user_input.send_keys(bpm_user)

    psw_input = wait.until(ec.presence_of_element_located((By.NAME, "pwd")))
    psw_input.send_keys(bpm_password)

    submit_button = wait.until(
        ec.presence_of_element_located((By.NAME, "submit"))
    )
    submit_button.click()


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in .env"
    return value


def main() -> None:
    project_folder = os.path.dirname(os.path.dirname(__file__))

    data_folder = os.path.join(project_folder, "data")
    os.makedirs(data_folder, exist_ok=True)

    report_folder = os.path.join(data_folder, "reports")
    os.makedirs(report_folder, exist_ok=True)

    download_folder = os.path.join(data_folder, "downloads")
    os.makedirs(download_folder, exist_ok=True)

    today = datetime.now().strftime("%d.%m.%y")

    bpm_user = get_from_env("BPM_USER")
    bpm_password = get_from_env("BPM_PASSWORD")

    driver_path = os.path.join(project_folder, "chromedriver.exe")

    report_file_name = f"Отчет_командировки_{today}.xlsx"
    report_file_path = os.path.join(report_folder, report_file_name)

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

    driver = driver_init(
        executable_path=driver_path, download_folder=download_folder
    )
    wait = WebDriverWait(driver, 10)

    with driver:
        login(driver, wait, bpm_user=bpm_user, bpm_password=bpm_password)
        wait.until(
            ec.visibility_of_element_located(
                (By.CSS_SELECTOR, ".cp_menu_section_div_v.cp_menu_simple")
            )
        )
        driver.get(
            "https://bpmtest.kdb.kz/?s=rep_b&id=13635&reset_page=1&gid=739"
        )

        do_reports_exist = (
            len(driver.find_elements(By.CSS_SELECTOR, ".empty_notice_header"))
            == 0
        )

        if do_reports_exist:
            as_excel = wait.until(
                ec.visibility_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="tab_form_card"]/table/tbody/tr[1]/td/div[1]/table/tbody/tr/td/input[4]',
                    )
                )
            )
            as_excel.click()


if __name__ == "__main__":
    main()

