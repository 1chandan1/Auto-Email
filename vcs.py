import os
import sys
import socket
import requests
import datetime
import subprocess

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Constants
GITHUB_EXE_URL = 'https://raw.githubusercontent.com/ChandanHans/Auto-email/main/output/AutoEmail.exe'
GITHUB_PY_URL = 'https://raw.githubusercontent.com/ChandanHans/Auto-email/main/auto_email.py'
REPO_API_URL = 'https://api.github.com/repos/ChandanHans/Auto-email/commits?path=output/AutoEmail.exe'
LOCAL_DATE_PATH = resource_path("date.txt")  # Path to the local version file
UPDATER_EXE_PATH = resource_path("updater.exe")  # The path to your updater executable
MAIN_FILE = 'auto_email.py'
EXE_PATH = sys.executable

def get_local_version_date():
    """Read the version date from the local version file."""
    with open(LOCAL_DATE_PATH, 'r') as file:
        return datetime.datetime.strptime(file.read().strip(), '%Y-%m-%dT%H:%M:%SZ')

def get_remote_version_date():
    """Fetch the latest commit date of the version file from the GitHub repository."""
    response = requests.get(REPO_API_URL)
    commits = response.json()
    if not commits:
        print("No commits found for the version file.")
        return None
    return datetime.datetime.strptime(commits[0]['commit']['committer']['date'], '%Y-%m-%dT%H:%M:%SZ')

def get_local_script_content():
    with open(MAIN_FILE, 'r', encoding="utf-8") as file:
        content = file.read()
    return content

def get_remote_script_content():
    response = requests.get(GITHUB_PY_URL)
    content = response.text
    return content

def write_new_script(new_script):
    with open(MAIN_FILE, 'w', encoding="utf-8") as file:
        file.write(new_script)
        
def is_my_machine():
    my_machine_list = ['ASUS-CHANDAN']
    return socket.gethostname() in my_machine_list
    
def check_for_updates():
    """Check if an update is available based on the latest commit date."""
    print("Checking for updates...")
    if getattr(sys, "frozen", False): 
        try:
            local_version_date = get_local_version_date()
        except:
            return
        remote_version_date = get_remote_version_date()

        if remote_version_date is None:
            return
        # Calculate the difference in time
        time_difference = remote_version_date - local_version_date
        # Check if the difference is greater than 2 minutes
        if time_difference > datetime.timedelta(minutes=2):
            try:
                subprocess.Popen([UPDATER_EXE_PATH, EXE_PATH, GITHUB_EXE_URL])
            except:
                input("ERROR : Contact Chandan")
            sys.exit()
    else:
        local_script_content = get_local_script_content()
        remote_script_content = get_remote_script_content()
        if not ((local_script_content == remote_script_content) or is_my_machine()):
            write_new_script(remote_script_content)
            print("Script Updated")
            input("Please close this app and restart it again")
            sys.exit()