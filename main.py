from vcs import check_for_updates

check_for_updates()

import locale

from client_email import client_email
from facturation import facturation
from google_services import GoogleServices
from notary_email import notary_email
from utils import *


def start(user: GoogleServices):
    user.print_details()
    print("1. Notary Email")
    print("2. Client Email")
    print("3. Facturation")
    print("\nEnter your choice (1/2/3): ")
    while True:
        choice = getch()
        print(choice)
        if choice == "1":
            print("\nLoading...")
            notary_email(user)
            break
        elif choice == "2":
            print("\nLoading...")
            client_email(user)
            break
        elif choice == "3":
            print("\nLoading...")
            facturation(user)
            break
    sleep(0.1)
    start(user)


def main():
    locale.setlocale(locale.LC_TIME, "fr_FR")
    user = GoogleServices()
    start(user)


if __name__ == "__main__":
    main()
