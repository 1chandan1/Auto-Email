from datetime import datetime
import re
import fitz  # PyMuPDF
from PIL import Image
import gspread
import io
from time import sleep

from google_services import GoogleServices
from constants import *
from image_processing import ask_date
from utils import *

def assign_case(user : GoogleServices):
    notary_worksheet = user.gc.open_by_key(NOTARY_SHEET_KEY).get_worksheet(0)
    invoice_worksheet = user.gc.open_by_key(INVOICE_SHEET_KEY).get_worksheet(0)
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
    notary_name = all_notary_data[notary_index][0]
    verify_folder_id = all_notary_data[notary_index][5]
    submit_folder_id = all_notary_data[notary_index][2]
    
    user.print_details()
    print()
    print_center(notary_name)
    print("\n")
    
    folders_to_verify = user.get_target_folders(verify_folder_id)
    try:
        for folder in folders_to_verify:
            folder_name = str(folder["name"])
            folder_id = str(folder["id"])
            print("------------------------------------------------------------")
            print(f"\n\n ------> {folder_name}")
            case_name = folder_name.split("(")[0].strip()
            collab = get_collab(folder_name)
            if any(word.isupper() for word in case_name.split()):
                all_good = verify_folder(user, folder_id)
                if all_good:
                    user.move_folder(folder_id,verify_folder_id, submit_folder_id)
                    update_invoice_sheet(invoice_worksheet, case_name, collab, notary_name)
                    print("All Good")
                    
            else:
                print("Problem in Folder name")
            print("------------------------------------------------------------")
    except Exception as e:
        print(e)
    input("\n\nPress Enter to Exit :")
    

def get_default_collab():
    print("\nCollab :")
    for index,value in enumerate(COLLAB_LIST,start=1):
        print(f"{index}. {value}")
    print()
    print("Enter your choice : ")
    while True:
        choice = getch()
        print(choice)
        try:
            if 0 < int(choice) <= len(COLLAB_LIST):
                return COLLAB_LIST[int(choice)-1]
        except:
            pass
        sleep(0.1)


def verify_folder(user : GoogleServices, folder_id):
    death_proof_file = user.get_file_by_name(folder_id, "acte de décè")
    mandat_file = user.get_file_by_name(folder_id, "mandat")
    
    if not death_proof_file:
        print("No  ---  Acte de décès")
        return False
    if not mandat_file:
        print("No  ---  MANDAT")
        return False
    
    if user.get_file_size(death_proof_file["id"]) > 4:
        print("Acte de décès > 4mb")
        return False
    
    if user.get_file_size(mandat_file["id"]) > 4:
        print("MANDAT > 4mb")
        return False
    
    date_dict  = get_dob_dod(user, death_proof_file["id"])
    if not date_dict:
        print("problem in  ---  dob_dod.json")
        return False
    print(date_dict)
    user.upload_json(folder_id,date_dict,"dob_dod.json")
    return True
    
def get_collab(folder_path):
    match = re.search(r'\((.*?)\)', folder_path)

    if match:
        value = match.group(1)
        return value
    return None
    
def update_invoice_sheet(worksheet : gspread.Worksheet, name, collab, notary_name):
    existing_case_name = [str(value).strip() for value in worksheet.col_values(2)]
    next_row = len(existing_case_name) + 1
    if not name in existing_case_name:
        worksheet.insert_row([None,name,None,None,collab, notary_name],next_row)

def get_dob_dod(user : GoogleServices, death_proof_id):
    date_dict = {}
    death_proof_pdf = user.download_file(death_proof_id)
    if death_proof_pdf:
        img = get_death_proof_img(death_proof_pdf)
        result = ask_date(img)
        if result:
            date_dict = fix_dob_dod(result, death_proof_id)
    return date_dict
    
    
def fix_dob_dod(date_dict, file_id):
    if 'DOB' not in date_dict or 'DOD' not in date_dict:
        print("\nThere is a problem.")
        print(f"This is the death certificate : https://drive.google.com/file/d/{file_id}.")
        print("Manually Enter the DOB and DOD")
        # Prompt user to input DOB and DOD
        date_dict = {
            "DOB": input("Enter DOB (DD/MM/YYYY): "),
            "DOD": input("Enter DOD (DD/MM/YYYY): ")
        }

    # Check if the dates are valid
    if is_valid_date(date_dict["DOB"]) and is_valid_date(date_dict["DOD"]):
        return date_dict
    else:
        # If dates are invalid, reset the dictionary and reattempt the process
        return fix_dob_dod({}, file_id)

def is_valid_date(date_text):
    try:
        if datetime.strptime(date_text, "%d/%m/%Y") and re.match(r'^\d{2}/\d{2}/\d{4}$', date_text):
            return True
        else:
            return False
    except ValueError:
        return False
    
def get_death_proof_img(pdf_file : io.BytesIO, resolution=150):
    doc = fitz.open("pdf", pdf_file.read())
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(resolution / 72, resolution / 72))
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return img