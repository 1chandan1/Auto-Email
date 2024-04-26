from time import sleep
import easygui
from google_services import GoogleServices
from constants import *
from utils import *

def upload_case(user : GoogleServices):
    notary_worksheet = user.gc.open_by_key(NOTARY_SHEET_KEY).get_worksheet(0)
    all_notary_data = notary_worksheet.get_values()[1:]
    
    user.print_details()
    choice_list = []
    for num, notary_data in enumerate(all_notary_data, start=1):
        print(f"{num}. {notary_data[0]}")
        choice_list.append(str(num))
    print("\nEnter your choice ( 0 : quit ) : ")
    while True:
        choice = getch()
        if choice in choice_list:
            break
        elif choice == "0":
            return
        sleep(0.1)
    
    notary_index = int(choice) - 1
    notary_folder_id = all_notary_data[notary_index][2]
    
    user.print_details()
    print()
    print_center(all_notary_data[notary_index][0])
    print("\n")
    input("Press Enter to select the root folder : ")
    
    root_folder_path = easygui.diropenbox()
    print("Uploading...\n\n")
    user.upload_folder(root_folder_path, notary_folder_id)
        
# def start_uploading(user, folder_id):



    
if __name__ == "__main__":
    upload_case(GoogleServices())