import re
import json
import os
import shutil
import sys
from getch import getch as pygetch
from time import sleep
from dotenv import load_dotenv


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


load_dotenv(dotenv_path=resource_path(".env"))
