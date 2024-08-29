import os
import shutil
from datetime import datetime
import pandas as pd

PATH = r"C:\Users\robotX4\Desktop\entering_data_on _orders\Командировки"

downloads_folder = r"C:\Users\robotX4\Downloads"
work_folder = PATH

today = datetime.now().strftime("%d.%m.%y")
current_day_report_path = os.path.join(work_folder, f"orders_{today}.xlsx")

for download_file in os.listdir(downloads_folder):
    if download_file.startswith("rep") and download_file.endswith(".xlsx"):
        shutil.copyfile(
            os.path.join(downloads_folder, download_file), current_day_report_path
        )

df = pd.read_excel(current_day_report_path, skiprows=3)
df = df.rename(
    columns={
        "Имя сотрудника": "employee_fullname",
        "Номер приказа": "order_number",
        "Дата подписания": "sign_date",
        "Дата начала": "start_date",
        "Дата окончания": "end_date",
        "Место командирования": "trip_place",
        "Цель командировки": "trip_target",
        "Номер основного приказа": "main_order_number",
        "Дата начала основного приказа": "main_order_start_date",
        "Имя замещающего сотрудника": "deputy_fullname",
    }
)

df = df.dropna(subset=["employee_fullname"])
df = df.dropna(subset=["sign_date"])

assert len(df.columns) == 10, "Странное кол-во колонок"

df.loc[:, "employee_names"] = df["employee_fullname"].str.split()
df.loc[:, "deputy_names"] = df["deputy_fullname"].str.split()
df.loc[:, "sign_date"] = df["sign_date"].str.replace(".", "")
df.loc[:, "start_date"] = df["start_date"].str.replace(".", "")
df.loc[:, "end_date"] = df["end_date"].str.replace(".", "")

df.to_json(
    os.path.join(work_folder, f"orders_{today}.json"),
    orient="records",
    force_ascii=False,
    indent=2,
)
