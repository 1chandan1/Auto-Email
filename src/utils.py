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


def get_client_secret():
    client_secret = os.environ["CLIENT_SECRET"]
    return json.loads(client_secret)


load_dotenv(dotenv_path=resource_path(".env"))
