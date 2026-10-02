import json, shutil

from pathlib import Path
from datetime import datetime as dt, timedelta

now = dt.now()
last_month = now.replace(day=1) - timedelta(days=1)
double_last_month = last_month.replace(day=1) - timedelta(days=1)


def format_date(date):
    """

    :type date: dt.date
    """
    return date.strftime("%Y-%m")


def folder_format(date):
    """

    :type date: dt.date
    """
    return date.strftime("%Y-%m-%d")


def load_client_info(client_name, clients_file="clients.json"):
    with open(clients_file, "r") as clients:
        data = json.load(clients)
    return data.get(client_name)


def make_filename(filepath: Path, client_name: str) -> Path | None:
    filename = filepath.name.removesuffix(".xlsx")
    words = filename.lower().split(" ")
    formatted_date = last_month.strftime("%Y-%m")
    if "detail" in words:
        new_path = filepath.with_name(f"{formatted_date} {client_name} P&L Detail{filepath.suffix}")
        print(new_path)
        return new_path
    elif "balance" in words:
        new_path = filepath.with_name(f"{formatted_date} {client_name} Balance Sheet{filepath.suffix}")
        print(new_path)
        return new_path
    elif "profit and loss" in words and "detail" not in words:
        new_path = filepath.with_name(f"{formatted_date} {client_name} P&L by Month{filepath.suffix}")
        print(new_path)
        return new_path
    return None


def move_to_client_folder(client_folder: Path, staged_file):
    return staged_file.rename(f"{client_folder}\\{staged_file.name}")


def cleanup_folder(client_folder: Path):
    prefix = format_date(double_last_month)
    new_folder = client_folder / f"{double_last_month.month}. {double_last_month.strftime('%B')}"

    for file in list(client_folder.iterdir()):
        if file.is_file() and file.name.startswith(prefix):
            new_folder.mkdir(parents=True, exist_ok=True)
            shutil.move(file, new_folder / file.name)
