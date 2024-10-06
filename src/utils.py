import re
import json
import os
import shutil
import sys
from getch import getch as pygetch
from time import sleep
from dotenv import load_dotenv
from unidecode import unidecode

def print_center(text):
    terminal_width = shutil.get_terminal_size().columns
    padding = (terminal_width - len(text)) // 2
    print(" " * padding + text)


def clear_display():
    os.system("cls" if os.name == "nt" else "clear")


def countdown(text: str, t: int):
    while t >= 0:
        print(f"{text} : {t} sec", end="\r")
        sleep(1)
        t -= 1
    print()


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    return os.path.join(base_path, relative_path)


def getch():
    if sys.platform == "win32":
        return pygetch().decode().lower()
    else:
        return pygetch().lower()

def extract_row_number(range_text):
    match = re.search(r":\D*(\d+)", range_text)
    if match:
        return int(match.group(1))
    else:
        return None

def is_valid_email(email):
    # Regular expression pattern for validating an email
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    # Match the pattern with the email
    if re.match(pattern, email):
        return True
    else:
        return False

def get_id_from_url(url):
    if url:
        match = re.search(r'[-\w]{25,}', url)
        return match.group(0) if match else None
    return None

def get_client_secret():
    client_secret = os.environ["CLIENT_SECRET"]
    return json.loads(client_secret)


def get_image_url_by_name(name, all_image_data):
    name = re.sub(r"[ ,\-\n]", "",unidecode(name).strip().lower())
    for i in all_image_data:
        image_name = re.sub(r"[ ,\-\n]", "",unidecode(i[0]).strip().lower())
        if name and image_name and name in image_name:
            return i[1]
    return None


def find_rows_by_email(data, email):
    return [
        index
        for index, sublist in enumerate(data, start=1)
        if email in sublist
    ]

def get_fname_lname(full_name: str):
    words = full_name.strip().split()
    lname = (
        " ".join([word for word in words if word.isupper() or word.lower() == "de"])
    ).upper()
    fname = (
        " ".join([word for word in words if not word.isupper() and word != "de"])
    )
    return fname, lname

load_dotenv(dotenv_path=resource_path(".env"))
