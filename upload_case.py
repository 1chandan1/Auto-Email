import re
import fitz  # PyMuPDF
from PIL import Image
import gspread
import io
import tkinter as tk
from time import sleep
from tkinter import filedialog

import pandas as pd
from google_services import GoogleServices
from constants import *
from image_processing import get_dob_dod
from utils import *

def upload_case(user : GoogleServices):
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
    notary_folder_id = all_notary_data[notary_index][2]
    
    default_collab = get_default_collab()
    
    user.print_details()
    print()
    print_center(notary_name)
    print("\n")
    input("Press Enter to select the root folder : ")
    root_folder_path = get_root_folder_path()
    try:
        for folder_name in os.listdir(root_folder_path):
            print("------------------------------------------------------------")
            print(f"\n\n ------> {folder_name}")
            drive_folder_name = folder_name
            case_name = folder_name.split("(")[0].strip()
            collab = get_collab(folder_name)
            if not collab:
                collab = default_collab
                drive_folder_name = folder_name + f" ({collab})"
            
            if any(word.isupper() for word in case_name.split()):
                folder_path = os.path.join(root_folder_path,folder_name)
                
                date_dict = process_folder(folder_path)
                if date_dict:
                    print("Uploading...\n")
                    user.upload_folder(drive_folder_name, folder_path, notary_folder_id)
                    update_invoice_sheet(invoice_worksheet, case_name,date_dict, collab, notary_name)
                    shutil.rmtree(folder_path)
            else:
                print("Problem in Folder name")
            print("------------------------------------------------------------")
    except Exception as e:
        print(e)
    input("\n\nPress Enter to Exit :")
    
    
def get_root_folder_path():
    while True:
        root = tk.Tk()
        root.withdraw()  # Hide the Tkinter root window
        root.attributes('-topmost', True)
        path = filedialog.askdirectory()
        root.destroy()
        print(f"\nPath: {path}")
        print("\nPlease Confirm (y/n)")
        while True:
            choice = getch()
            print(choice)
            if choice == "y":
                return path
            elif choice == "n":
                break
            sleep(0.1)

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


def process_folder(folder_path):
    all_file_name = [str(file_name) for file_name in os.listdir(folder_path)]
    death_proof_file = next((element for element in all_file_name if "acte de décè" in element.lower()), None)
    mandat_file = next((element for element in all_file_name if "mandat" in element.lower()), None)
    
    if not death_proof_file:
        print("No  ---  Acte de décès")
        return None
    if not mandat_file:
        print("No  ---  MANDAT")
        return None
    death_proof_path = os.path.join(folder_path,death_proof_file)
    mandat_path = os.path.join(folder_path,mandat_file)
    
    if not is_goog_size(death_proof_path):
        print("Acte de décès > 4mb")
        return None
    
    if not is_goog_size(mandat_path):
        print("MANDAT > 4mb")
        return None
    
    date_dict  = save_dob_dod(folder_path,death_proof_file)
    if not date_dict:
        print("problem in  ---  dob_dod.json")
        return None
    return date_dict
    
def get_collab(folder_path):
    match = re.search(r'\((.*?)\)', folder_path)

    if match:
        value = match.group(1)
        return value
    return None
    
def update_invoice_sheet(worksheet : gspread.Worksheet, name,date_dict, collab, notary_name):
    existing_case_name = [str(value).strip() for value in worksheet.col_values(2)]
    next_row = len(existing_case_name) + 1
    if not name in existing_case_name:
        worksheet.insert_row([None,name,None,None,collab, notary_name],next_row)
        worksheet.update_acell(f"C{next_row}", date_dict["DOB"])
        worksheet.update_acell(f"D{next_row}", date_dict["DOD"])

def save_dob_dod(folder_path,death_proof_file):
    death_proof_path = os.path.join(folder_path,death_proof_file)
    date_path = os.path.join(folder_path,"dob_dod.json")
    try:
        with open(date_path,"r") as file:
            result = json.load(file)
    except:
        img = get_death_proof_img(death_proof_path)
        result = get_dob_dod(img)
        with open(date_path,"w") as file:
            json.dump(result, file, indent=4)
    
    if result["DOB"] and result["DOD"]:
        return result
    return None
    
def get_death_proof_img(pdf_path, resolution=300):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(resolution / 72, resolution / 72))
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return img
    

def is_goog_size(file_path):
    size = os.path.getsize(file_path)
    return size < 4 * 1024 * 1024
    