import requests
import datetime
import sys
import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# Constants
REPO_API_URL = 'https://api.github.com/repos/1chandan1/Auto-email/commits?path=output/AutoEmail.exe'
LOCAL_VERSION_PATH = resource_path("version.txt")  # Path to the local version file
EXE_PATH = sys.executable
UPDATER_EXE_PATH = resource_path("updater.exe")  # The path to your updater executable

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

def check_for_updates():
    """Check if an update is available based on the latest commit date."""
    print("Checking for updates...")
    local_version_date = get_local_version_date()
    remote_version_date = get_remote_version_date()

    if remote_version_date is None:
        return

    if local_version_date < remote_version_date:
        print(f"Update available. Local version date: {local_version_date}, Remote version date: {remote_version_date}")
        # Implement your logic to handle an available update here.
        # For example, initiate_update(new_exe_path)
    else:
        print("No update available. Running the latest version.")

