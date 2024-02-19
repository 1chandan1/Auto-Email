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
GITHUB_EXE_URL = 'https://raw.githubusercontent.com/ChandanHans/Auto-email/main/output/AutoEmail.exe'
REPO_API_URL = 'https://api.github.com/repos/ChandanHans/Auto-email/commits?path=output/AutoEmail.exe'
LOCAL_DATE_PATH = resource_path("date.txt")  # Path to the local version file

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

def check_for_updates():
    """Check if an update is available based on the latest commit date."""
    try:
        print("Checking for updates...")
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
            print(f"Update available")
            input(f"Please download the new Application : {GITHUB_EXE_URL}")
            sys.exit()
    except Exception as e:
        print(e)
        print("\n\n!! Error !!")
        input("Press Enter to EXIT : ")
        sys.exit()