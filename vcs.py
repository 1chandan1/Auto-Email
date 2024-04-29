import base64
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
GITHUB_EXE_URL = 'https://github.com/ChandanHans/Auto-email/raw/main/output/AutoEmail.exe'
REPO_API_URL = f"https://api.github.com/repos/ChandanHans/Auto-email/git/trees/main?recursive=1"
REPO_EXE_API_URL = 'https://api.github.com/repos/ChandanHans/Auto-email/commits?path=output/AutoEmail.exe'
LOCAL_DATE_PATH = resource_path("date.txt")
UPDATER_EXE_PATH = resource_path("updater.exe")
EXE_PATH = sys.executable

def get_local_version_date():
    """Read the version date from the local version file."""
    with open(LOCAL_DATE_PATH, 'r') as file:
        return datetime.datetime.strptime(file.read().strip(), '%Y-%m-%dT%H:%M:%SZ')

def get_remote_version_date():
    """Fetch the latest commit date of the version file from the GitHub repository."""
    response = requests.get(REPO_EXE_API_URL)
    commits = response.json()
    if not commits:
        print("No commits found for the version file.")
        return None
    return datetime.datetime.strptime(commits[0]['commit']['committer']['date'], '%Y-%m-%dT%H:%M:%SZ')
        
def is_my_machine():
    my_machine_list = ['ASUS-CHANDAN']
    return socket.gethostname() in my_machine_list
    
def update_local_files():
    updated = False
    response = requests.get(REPO_API_URL)
    if response.status_code != 200:
        print(f"Failed to fetch repository data: {response.status_code}")
        return

    files = response.json().get('tree', [])
    python_files = [file for file in files if file['path'].endswith('.py')]
    for file in python_files:
        file_url = file['url']
        download_response = requests.get(file_url)
        if download_response.status_code == 200:
            file_content_encoded = download_response.json()['content']
            file_content = base64.b64decode(file_content_encoded).decode('utf-8')  # Decode from base64
            file_path = os.path.join(os.getcwd(),file['path'])
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            local_content = ""
            try:
                with open(file_path, 'r', encoding="utf-8") as f:
                    local_content = f.read()
            except:
                pass
            if local_content != file_content:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(file_content)
                updated = True
                print(f"Updated: {file['path']}")
        else:
            print(f"Failed to download {file['path']}: {download_response.status_code}")
    return updated


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
        if not is_my_machine() and update_local_files():
            print("Script Updated")
            input("Please close this app and restart it again")
            sys.exit()
