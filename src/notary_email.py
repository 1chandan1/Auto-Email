import base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random
import re
from googleapiclient.errors import HttpError
from gspread.exceptions import SpreadsheetNotFound
from dateutil.relativedelta import relativedelta
import gspread
import pandas as pd

from time import sleep

from unidecode import unidecode
from src.constants import *
from src.google_services import GoogleServices
from src.utils import countdown, getch, print_center


def notary_email(user: GoogleServices):
    user.print_details()
    print_center("Notary Email")
    print("\n")
    while True:
        url = input("Target Google Sheet Link ( 0 : quit ) : ").strip()
        if url == "0":
            return
        try:
            spreadsheet = user.gc.open_by_url(url)
            break
        except SpreadsheetNotFound:
            print("The specified spreadsheet was not found.")
        except HttpError:
            print("\nNo Internet Connection\n")
        except Exception as e:
            print(f"ERROR")
    user.print_details()
    print_center("Notary Email")
    print("\n")
    print_center(f"Google Sheet : {spreadsheet.title}")
    print("\n")
    print("1. Send Emails")
    print("2. Change Google Sheet")
    print("q. Main menu")
    print("\nEnter your choice (1/2/q): ")
    while True:
        choice = getch()
        print(choice)
        if choice == "1":
            print("\nLoading...")
            send_notary_emails(user, spreadsheet)
            user.print_details()
            print_center("Notary Email")
            print()
            input("\n\nTask Completed\nPress Enter To Continue : ")
            break
        elif choice == "2":
            notary_email(user)
            break
        elif choice == "q":
            break
    sleep(0.1)


def send_notary_emails(user: GoogleServices, spreadsheet: gspread.Spreadsheet):
    annuraie_sheet = user.gc.open_by_key(ANNUAIRE_SHEET_KEY)
    annuraie_worksheet = annuraie_sheet.worksheet(ANNUAIRE_WORKSHEET_NAME)
    scheduling_worksheet = annuraie_sheet.worksheet(SCHEDULED_WORKSHEET_NAME)
    worksheet = spreadsheet.get_worksheet(0)
    sheet_data = worksheet.get_values()
    header = sheet_data[0]
    rows = sheet_data[1:]
    df = pd.DataFrame(rows, columns=header)
    df.index += 2
    status_col_index = df.columns.get_loc("Status") + 1
    comment_column_index = status_col_index + 1
    all_annuraie_data = annuraie_worksheet.get_all_values()
    all_scheduled_data = scheduling_worksheet.get_all_values()
    all_cases = get_all_cases(all_annuraie_data, all_scheduled_data)
    for index, row in df.iterrows():
        try:
            if row["Status"] == "à envoyer":
                notary_email : str = row["Email"].split("\n")[0].strip()
                person_full_name : str = row["Name"].strip()
                _, person_last_name = get_fname_lname(person_full_name)
                notary_full_name : str = re.sub(r"[ ,\-\n]", " ",unidecode(row["Notary"]).strip())
                notary_first_name, notary_last_name = get_fname_lname(notary_full_name)
                person_don  : str = row["Data Of Notoriety"].strip()

                if not all(
                    [
                        person_last_name.strip(),
                        notary_last_name.strip(),
                        person_don,
                        person_don != "N/A",
                        is_valid_email(notary_email),
                    ]
                ):
                    worksheet.update_cell(
                        index, comment_column_index, "Problem"
                    )
                    continue

                
                notary_sheet_index = find_row_by_name(
                    all_annuraie_data, notary_first_name, notary_last_name
                )
                
                
                if notary_sheet_index:
                    annuraie_worksheet.update_acell(f"I{notary_sheet_index}", notary_email)
                    all_annuraie_data[notary_sheet_index - 1][8] = notary_email
                    
                email_row_indices = find_rows_by_email(all_annuraie_data, notary_email)
                notary_status, contact_date, first_case = get_status_date_case(all_annuraie_data, email_row_indices)
                
                
                if not notary_sheet_index:  
                    new_row = [
                        notary_first_name,
                        notary_last_name,
                        "","","","",
                        row["Phone"],
                        "",
                        notary_email,
                        notary_status,
                        contact_date,"","","",
                    ]
                    annuraie_updated_row = annuraie_worksheet.append_row(
                        new_row, value_input_option="USER_ENTERED", table_range='A:A'
                    )
                    updated_range = annuraie_updated_row["updates"]["updatedRange"]
                    notary_sheet_index = extract_row_number(updated_range)
                    all_annuraie_data.append(new_row)
                    email_row_indices.append(notary_sheet_index)
                    worksheet.update_cell(
                        index, comment_column_index, "New Notary added"
                    )
                    
                if notary_sheet_index:
                    annuraie_worksheet.update_acell(f"J{notary_sheet_index}", notary_status)
                    annuraie_worksheet.update_acell(f"K{notary_sheet_index}", contact_date)
                    annuraie_worksheet.update_acell(f"N{notary_sheet_index}", first_case)
                    annuraie_worksheet.update_acell(f"Q{notary_sheet_index}", user.email)
                    all_annuraie_data[notary_sheet_index - 1][9] = notary_status
                    all_annuraie_data[notary_sheet_index - 1][10] = contact_date
                    all_annuraie_data[notary_sheet_index - 1][13] = first_case
            
                annuraie_sheet_row = all_annuraie_data[notary_sheet_index - 1]
                
                user.print_details()
                print_center("Notary Email")
                print()
                print_center(f"Google Sheet : {spreadsheet.title}")
                print()
                print_center("Sending All Emails\n\n")
                print(f"Index-File Row    :    {notary_sheet_index}")
                print(f"Contact Date      :    {contact_date}\n")
                print(f"Target Sheet Row  :    {index}\n")
                print(f"Person Name       :    {person_full_name}")
                print(f"Person Last Name  :    {person_last_name}\n")
                print(f"Notary Name       :    {notary_full_name}")
                print(f"Notary Last Name  :    {notary_last_name}\n")
                print(f"DON               :    {person_don}")
                print(f"To                :    {notary_email}\n")
                
                
                name_for_checking = re.sub(
                    r"[ ,\-\n]", "", unidecode(person_full_name).lower().strip()
                )
                if name_for_checking not in all_cases:
                    all_cases.append(name_for_checking)
                    if notary_status == "Not contacted":
                        # countdown("Sending Email in", random.randint(120, 180))
                        print("\nSending Email...")
                        status = None
                        for _ in range(3):
                            message = create_notary_message(
                                user,
                                notary_email,
                                person_full_name,
                                person_last_name,
                                notary_last_name,
                                person_don,
                            )
                            status = user.send_email(message)
                            if status:
                                break
                            countdown("Trying to send email again", 5)

                        if status:
                            worksheet.update_cell(index, status_col_index, "envoyé")
                            today_date = datetime.now().date().strftime("%d-%b-%Y")

                            annuraie_worksheet.update_acell(f"K{notary_sheet_index}", today_date)
                            annuraie_worksheet.update_acell(
                                f"J{notary_sheet_index}", "Contacted / pending answer"
                            )
                            annuraie_worksheet.update_acell(
                                f"N{notary_sheet_index}", person_full_name
                            )
                            annuraie_worksheet.update_acell(f"Q{ notary_sheet_index }", user.email)
                            all_annuraie_data[notary_sheet_index - 1][9] = "Contacted / pending answer"
                            all_annuraie_data[notary_sheet_index - 1][10] = today_date
                            all_annuraie_data[notary_sheet_index - 1][13] = person_full_name
                            print("\nSuccess")
                        else:
                            worksheet.update_cell(
                                index, comment_column_index, f"Error"
                            )
                    elif notary_status in (
                        "Contacted / pending answer",
                        "Cooperating",
                        "Collègue coopérant"
                    ):
                        print("Scheduling Email")
                        sleep(2)
                        previous_scheduled_date = annuraie_sheet_row[10]
                        new_date_text = None
                        if notary_status == "Cooperating":
                            for scheduled_data in all_scheduled_data[::-1]:
                                if notary_email in scheduled_data:
                                    previous_scheduled_date = scheduled_data[9]
                                    break
                            try:
                                new_date = datetime.strptime(
                                    previous_scheduled_date, "%d-%b-%Y"
                                ) + relativedelta(months=+2)
                            except:
                                new_date = datetime.now().date() + relativedelta(
                                    months=+2
                                )

                            if new_date.weekday() == 5:
                                new_date += relativedelta(days=2)
                            elif new_date.weekday() == 6:
                                new_date += relativedelta(days=1)

                            new_date_text = new_date.strftime("%d-%b-%Y")

                        new_schedule_row = [
                            notary_first_name,
                            notary_last_name,
                            None,
                            "Scheduled",
                            person_full_name,
                            person_don,
                            None,
                            user.email,
                            notary_email,
                            new_date_text,
                        ]
                        scheduling_worksheet.append_row(
                            new_schedule_row, value_input_option="USER_ENTERED",table_range='A:A'
                        )
                        all_scheduled_data.append(new_schedule_row)
                        worksheet.update_cell(
                            index,
                            comment_column_index,
                            f"Scheduled {new_date_text or ''}",
                        )
                        worksheet.update_cell(index, status_col_index, "draft")
                        print(f"Scheduled {new_date_text}")
                    elif notary_status == "Not cooperating":
                        worksheet.update_cell(
                            index, comment_column_index, "Not cooperating"
                        )
                else:
                    worksheet.update_cell(index, status_col_index, "draft")
                    worksheet.update_cell(
                        index, comment_column_index, f"Already scheduled"
                    )
                    print(f"Already scheduled")
                countdown("Next", 5)
        except Exception as e:
            print(e)
            countdown("Next", 5)


def get_all_cases(all_annuraie_data, all_scheduled_data):
    all_scheduled_cases = [
        re.sub(r"[ ,\-\n]", "", unidecode(row[4]).lower().strip())
        for row in all_scheduled_data
    ]
    all_annuraie_cases = [
        re.sub(r"[ ,\-\n]", "", unidecode(row[13]).lower().strip())
        for row in all_annuraie_data
    ]
    all_cases = list(set(all_annuraie_cases + all_scheduled_cases))
    return all_cases

def extract_row_number(range_text):
    match = re.search(r":\D*(\d+)", range_text)
    if match:
        return int(match.group(1))
    else:
        return None


def get_fname_lname(full_name: str):
    words = full_name.strip().split()
    lname = (
        " ".join([word for word in words if word.isupper() or word.lower() == "de"])
    ).upper()
    fname = (
        " ".join([word for word in words if not word.isupper() and word != "de"])
    )
    return fname, lname


def is_valid_email(email):
    # Regular expression pattern for validating an email
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    # Match the pattern with the email
    if re.match(pattern, email):
        return True
    else:
        return False


def find_row_by_name(annuraie_data: list[list], first_name: str, last_name: str):
    pattern = r"[ ,\-\n]"
    first_name = unidecode(re.sub(pattern, "", first_name)).lower()
    last_name = unidecode(re.sub(pattern, "", last_name)).lower()
    for index, row in enumerate(annuraie_data, start=1):
        if (
            first_name == unidecode(re.sub(pattern, "", row[0])).lower()
            and last_name == unidecode(re.sub(pattern, "", row[1])).lower()
        ):
            return index
    return None

def find_rows_by_email(data, email):
    return [
        index
        for index, sublist in enumerate(data, start=1)
        if email in sublist
    ]

def get_status_date_case(data, row_indices):
    if row_indices:
        first_index = row_indices[0] - 1
        all_status = [ data[index - 1][9] for index in row_indices ] 
        all_date = [ data[index - 1][10] for index in row_indices ] 
        all_case = [ data[index - 1][13] for index in row_indices ] 
        if "Cooperating" in all_status:
            return "Collègue coopérant", max(all_date), max(all_case)
        
        if "Not cooperating" in all_status:
            return "Not cooperating", None, None
        return data[first_index][9], max(all_date), max(all_case)
    else:
        return "Not contacted", None, None


def create_notary_message(
    user: GoogleServices,
    to: str,
    person_full_name: str,
    person_last_name: str,
    notary_last_name: str,
    person_don: str,
):
    message = MIMEMultipart()
    message["From"] = f"Klero Genealogie <{user.email}>"
    message["To"] = to
    message["Subject"] = f"Succession {person_last_name} - Demande de mise en relation"

    with open(NOTARY_EMAIL_TEMPLATE_PATH, "r", encoding="utf-8") as file:
        html_template = file.read()
    message_html = html_template.format(
        notary_last_name=notary_last_name,
        person_full_name=person_full_name,
        person_don=person_don,
    )
    message.attach(MIMEText(message_html + user.signature, "html"))

    return {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")}
