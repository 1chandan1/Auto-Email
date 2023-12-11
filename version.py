import subprocess
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
GITHUB_EXE_URL = 'https://github.com/1chandan1/Auto-email/releases/download/EXE/AutoEmail.exe'
RELEASES_API_URL = 'https://api.github.com/repos/1chandan1/Auto-email/releases/tags/EXE'
LOCAL_VERSION_PATH = resource_path("version.txt")  # Path to the local version file
UPDATER_EXE_PATH = resource_path("updater.exe")  # The path to your updater executable
EXE_PATH = sys.executable

def get_local_version_date():
    """Read the version date from the local version file."""
    with open(LOCAL_VERSION_PATH, 'r') as file:
        return datetime.datetime.strptime(file.read().strip(), '%Y-%m-%dT%H:%M:%SZ')

def get_remote_version_date():
    """Fetch the release date of AutoEmail.exe from the GitHub repository releases."""
    response = requests.get(RELEASES_API_URL)
    release_data = response.json()
    release_date_str = release_data['published_at']  # Get the publication date of the release
    return datetime.datetime.strptime(release_date_str, '%Y-%m-%dT%H:%M:%SZ')


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
        subprocess.Popen([UPDATER_EXE_PATH, EXE_PATH, GITHUB_EXE_URL])
        sys.exit()
