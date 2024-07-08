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

                all_row_with_same_email = [
                    index
                    for index, sublist in enumerate(all_annuraie_data, start=1)
                    if notary_email in sublist
                ]
                notary_sheet_index, annuraie_sheet_row = get_row_by_name(
                    all_annuraie_data, notary_first_name, notary_last_name
                )
                if not notary_sheet_index:
                    temp_status = "Not contacted"
                    temp_date = None
                    temp_first_case = None
                    if all_row_with_same_email:
                        try:
                            temp_row = annuraie_worksheet.row_values(
                                all_row_with_same_email[0]
                            )
                            temp_status = temp_row[10]
                            temp_date = temp_row[11]
                            temp_first_case = temp_row[14]
                        except:
                            pass
                    annuraie_sheet_row = [
                        None,
                        notary_first_name,
                        notary_last_name,
                        None,
                        None,
                        None,
                        None,
                        row["Phone"],
                        None,
                        notary_email,
                        temp_status,
                        temp_date,
                        None,
                        None,
                        temp_first_case
                    ]
                    annuraie_updated_row = annuraie_worksheet.append_row(
                        annuraie_sheet_row, value_input_option="USER_ENTERED", table_range='A:A'
                    )
                    all_annuraie_data.append(annuraie_sheet_row)
                    updated_range = annuraie_updated_row["updates"]["updatedRange"]
                    notary_sheet_index = extract_row_number(updated_range)
                    worksheet.update_cell(
                        index, comment_column_index, "New Notary added"
                    )
                    all_row_with_same_email.append(notary_sheet_index)
                
                notary_status = annuraie_sheet_row[10]
                
                if notary_status == "Not cooperating":
                    worksheet.update_cell(
                        index, comment_column_index, "Not cooperating"
                    )
                    continue
                contact_date = annuraie_sheet_row[11]
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
                annuraie_worksheet.update_acell(f"J{notary_sheet_index}", notary_email)
                
                
                name_for_checking = re.sub(
                    r"[ ,\-\n]", "", unidecode(person_full_name).lower().strip()
                )
                if name_for_checking not in all_cases:
                    all_cases.append(name_for_checking)
                    if notary_status == "Not contacted":
                        countdown("Sending Email in", random.randint(120, 180))
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
                            for row_index in all_row_with_same_email:
                                annuraie_worksheet.update_acell(f"L{row_index}", today_date)
                                annuraie_worksheet.update_acell(
                                    f"K{row_index}", "Contacted / pending answer"
                                )
                                annuraie_worksheet.update_acell(
                                    f"O{row_index}", person_full_name
                                )
                                annuraie_worksheet.update_acell(f"R{row_index}", user.email)
                                all_annuraie_data[row_index-1][10] = "Contacted / pending answer"
                                all_annuraie_data[row_index-1][11] = today_date
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
                        previous_scheduled_date = annuraie_sheet_row[11]
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
        re.sub(r"[ ,\-\n]", "", unidecode(row[14]).lower().strip())
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


def get_row_by_name(annuraie_data: list[list], first_name: str, last_name: str):
    pattern = r"[ ,\-\n]"
    first_name = unidecode(re.sub(pattern, "", first_name)).lower()
    last_name = unidecode(re.sub(pattern, "", last_name)).lower()
    for index, row in enumerate(annuraie_data, start=1):
        if (
            first_name == unidecode(re.sub(pattern, "", row[1])).lower()
            and last_name == unidecode(re.sub(pattern, "", row[2])).lower()
        ):
            return index, row
    return None, None


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
