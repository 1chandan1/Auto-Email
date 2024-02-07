import subprocess
import requests
import datetime
import sys
import os

import tqdm

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Constants
GITHUB_EXE_URL = 'https://raw.githubusercontent.com/ChandanHans/Auto-email/main/output/AutoEmail.exe'
REPO_API_URL = 'https://api.github.com/repos/ChandanHans/Auto-email/commits?path=output/AutoEmail.exe'
LOCAL_VERSION_PATH = resource_path("version.txt")  # Path to the local version file
EXE_PATH = sys.executable
UPDATER_EXE_PATH = resource_path("updater.bat")  # The path to your updater executable

def get_local_version_date():
    """Read the version date from the local version file."""
    with open(LOCAL_VERSION_PATH, 'r') as file:
        return datetime.datetime.strptime(file.read().strip(), '%Y-%m-%dT%H:%M:%SZ')

def get_remote_version_date():
    """Fetch the latest commit date of the version file from the GitHub repository."""
    response = requests.get(REPO_API_URL)
    commits = response.json()
    if not commits:
        print("No commits found for the version file.")
        return None
    return datetime.datetime.strptime(commits[0]['commit']['committer']['date'], '%Y-%m-%dT%H:%M:%SZ')

def download_new_version(file_name: str , url):
    """Download the new version and save it as a new file."""
    print("Downloading")
    # Send a request to get the file
    response = requests.get(url, stream=True)
    # Get the total file size in bytes.
    total_size = int(response.headers.get('content-length', 0))
    # Replace the dot in the filename to create a new filename
    new_exe = file_name.replace(".", "_new.")
    # Open the file in binary write mode
    with open(new_exe, "wb") as file:
        # Download the file with progress bar
        for data in tqdm(response.iter_content(chunk_size=1024), total=total_size//1024, unit='KB'):
            file.write(data)
    return new_exe

def check_for_updates():
    """Check if an update is available based on the latest commit date."""
    print("Checking for updates...")
    local_version_date = get_local_version_date()
    remote_version_date = get_remote_version_date()

    if remote_version_date is None:
        return
    # Calculate the difference in time
    time_difference = remote_version_date - local_version_date
    # Check if the difference is greater than 2 minutes
    if time_difference > datetime.timedelta(minutes=2):
        new_exe = download_new_version(EXE_PATH,GITHUB_EXE_URL)
        subprocess.Popen([f'del "{EXE_PATH}"'])
        subprocess.Popen([f'rename "{new_exe}" "{EXE_PATH}"'])
        sys.exit()
