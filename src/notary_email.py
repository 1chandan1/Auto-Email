import base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import io
import random
import re
import sys
from googleapiclient.errors import HttpError
from gspread.exceptions import SpreadsheetNotFound
from dateutil.relativedelta import relativedelta
import gspread
import pandas as pd

from time import sleep

from unidecode import unidecode
from src.constants import *
from src.google_services import GoogleServices
from src.utils import *

import locale
locale.setlocale(locale.LC_TIME, "fr_FR")


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
            print("\n\nTask Completed")
            countdown("Exit", 5)
            sys.exit()
        elif choice == "2":
            notary_email(user)
            break
        elif choice == "q":
            break
    sleep(0.1)
 
 
def send_notary_emails(user: GoogleServices, spreadsheet: gspread.Spreadsheet):
    annuraie_sheet = user.gc.open_by_key(ANNUAIRE_SHEET_KEY)
    death_certificates_sheet = user.gc.open_by_key(DEATH_CERTIFICATES_SHEET_KEY)
    death_certificates_worksheet = death_certificates_sheet.get_worksheet_by_id(0)
    all_image_data = death_certificates_worksheet.get_all_values()
    annuraie_worksheet = annuraie_sheet.worksheet(ANNUAIRE_WORKSHEET_NAME)
    scheduling_worksheet = annuraie_sheet.worksheet(SCHEDULED_WORKSHEET_NAME)
    worksheet = spreadsheet.get_worksheet(0)
    sheet_data = worksheet.get_values()
    header = sheet_data[0]
    rows = sheet_data[1:]
    df = pd.DataFrame(rows, columns=header)
    df.index += 2
    status_col_index = df.columns.get_loc("Status") + 1
    comment_column_index = status_col_index + 2
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
                
                    
                email_row_indices = find_rows_by_email(all_annuraie_data, notary_email)
                notary_status = get_status(all_annuraie_data, email_row_indices)
                
                
                if not notary_sheet_index:  
                    new_row = [
                        notary_first_name,
                        notary_last_name,
                        "","","","",
                        row["Phone"],"",
                        notary_email,
                        notary_status,
                        "","","","","",
                        "",user.email
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
                    
            
                annuraie_sheet_row = all_annuraie_data[notary_sheet_index - 1]
                
                user.print_details()
                print_center("Notary Email")
                print()
                print_center(f"Google Sheet : {spreadsheet.title}")
                print()
                print_center("Sending All Emails\n\n")
                print(f"Index-File Row    :    {notary_sheet_index}") 
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
                
                image_data = image_name = None
                image_link = get_image_url_by_name(name_for_checking, all_image_data)
                image_id = get_id_from_url(image_link)
                if image_id:
                    image_data, image_name = user.download_file(image_id)
                    
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
                                image_data,
                                image_name
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
                                f"N{notary_sheet_index}", person_full_name
                            )
                            annuraie_worksheet.update_acell(f"Q{ notary_sheet_index }", user.email)
                            all_annuraie_data[notary_sheet_index - 1][10] = today_date
                            all_annuraie_data[notary_sheet_index - 1][13] = person_full_name
                            for i in email_row_indices:
                                annuraie_worksheet.update_acell(f"J{i}", "Contacted / pending answer"
                                )
                                all_annuraie_data[i - 1][9] = "Contacted / pending answer"
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

                        # Define your logic to calculate new_date
                        if notary_status == "Cooperating" or notary_status == "Collègue coopérant":
                            # List to store all previous scheduled dates
                            previous_dates = []
                            try:
                                previous_dates.append(datetime.strptime(previous_scheduled_date, "%d-%b-%Y").date())
                            except:
                                pass
                            
                            for scheduled_data in all_scheduled_data:
                                if notary_email in scheduled_data:
                                    try:
                                        # Extract the date and append to the list
                                        previous_scheduled_date = datetime.strptime(scheduled_data[9], "%d-%b-%Y").date()
                                        previous_dates.append(previous_scheduled_date)
                                    except:
                                        # Handle any issues with date parsing, skip if invalid
                                        continue

                            # Get the maximum date from the list of previous dates if available
                            if previous_dates:
                                max_previous_date = max(previous_dates)
                                new_date = max_previous_date + relativedelta(months=+2)
                            else:
                                # Fallback if no previous dates are found
                                new_date = datetime.now().date() + relativedelta(months=+2)

                            # Check if new_date falls on a weekend or holiday, and adjust accordingly
                            while new_date.weekday() in (5, 6) or new_date in HOLIDAY_DATES:
                                if new_date.weekday() == 5:  # Saturday
                                    new_date += relativedelta(days=2)
                                elif new_date.weekday() == 6:  # Sunday
                                    new_date += relativedelta(days=1)
                                elif new_date in HOLIDAY_DATES:  # Holiday
                                    new_date += relativedelta(days=1)
                            
                            # Convert new_date to the required format
                            new_date_text = new_date.strftime("%d-%b-%Y")

                        new_schedule_row = [
                            notary_first_name,
                            notary_last_name,
                            None,
                            "Scheduled" if new_date_text else None,
                            person_full_name,
                            person_don,
                            user.email,
                            None,
                            new_date_text,
                            None,
                            None,
                            None,
                            image_link
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

def get_status(data, row_indices):
    if row_indices:
        all_status = [ data[index - 1][9] for index in row_indices ] 
        if "Cooperating" in all_status or "Collègue coopérant" in all_status:
            return "Collègue coopérant"
        
        elif "Not cooperating" in all_status:
            return "Not cooperating"
        
        elif "Contacted / pending answer" in all_status:
            return "Contacted / pending answer"
        
        else:
            return "Not contacted"
    else:
        return "Not contacted"


def create_notary_message(
    user: GoogleServices,
    to: str,
    person_full_name: str,
    person_last_name: str,
    notary_last_name: str,
    person_don: str,
    attachment: io.BytesIO = None,  # attachment is io.BytesIO
    attachment_filename: str = None,  # Provide the filename separately
):
    message = MIMEMultipart()
    message["From"] = f"Klero Genealogie <{user.email}>"
    message["To"] = to
    message["Subject"] = f"Succession {person_last_name} - Demande de mise en relation"
    
    if attachment and attachment_filename:
        template_path = NOTARY_EMAIL_TEMPLATE2_PATH
    else:
        template_path = NOTARY_EMAIL_TEMPLATE1_PATH

    with open(template_path, "r", encoding="utf-8") as file:
        html_template = file.read()
        message_html = html_template.format(
        notary_last_name=notary_last_name,
        person_full_name=person_full_name,
        person_don=person_don
    )
    
    message.attach(MIMEText(message_html + user.signature, "html"))
    
    # If an attachment is provided, attach it to the email
    if attachment and attachment_filename:
        # Create the attachment using MIMEApplication
        part = MIMEApplication(attachment.read(), Name=attachment_filename)

        # Add the header for the attachment
        part.add_header(
            "Content-Disposition", f"attachment; filename={attachment_filename}"
        )

        # Attach the file to the email
        message.attach(part)
    
    return {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")}

def get_image_url_by_name(name, all_image_data):
    name = re.sub(r"[ ,\-\n]", "",unidecode(name).strip().lower())
    for i in all_image_data:
        image_name = re.sub(r"[ ,\-\n]", "",unidecode(i[0]).strip().lower())
        if name and image_name and name in image_name:
            return i[1]
    return None

holidayDates = [
    # 2024
    "01-janv.-2024", "29-mars-2024", "01-avr.-2024", "01-mai-2024", "08-mai-2024", "09-mai-2024", "19-mai-2024",
    "20-mai-2024", "14-juil.-2024", "15-août-2024", "01-nov.-2024", "11-nov.-2024", "25-déc.-2024", "26-déc.-2024",

    # 2025
    "01-janv.-2025", "18-avr.-2025", "21-avr.-2025", "01-mai-2025", "08-mai-2025", "29-mai-2025", "08-juin-2025",
    "09-juin-2025", "14-juil.-2025", "15-août-2025", "01-nov.-2025", "11-nov.-2025", "25-déc.-2025", "26-déc.-2025",

    # 2026
    "01-janv.-2026", "03-avr.-2026", "06-avr.-2026", "01-mai-2026", "08-mai-2026", "14-mai-2026", "24-mai-2026",
    "25-mai-2026", "14-juil.-2026", "15-août-2026", "01-nov.-2026", "11-nov.-2026", "25-déc.-2026", "26-déc.-2026",

    # 2027
    "01-janv.-2027", "26-mars-2027", "29-mars-2027", "01-mai-2027", "08-mai-2027", "06-mai-2027", "16-mai-2027",
    "17-mai-2027", "14-juil.-2027", "15-août-2027", "01-nov.-2027", "11-nov.-2027", "25-déc.-2027", "26-déc.-2027",

    # 2028
    "01-janv.-2028", "14-avr.-2028", "17-avr.-2028", "01-mai-2028", "08-mai-2028", "25-mai-2028", "04-juin-2028",
    "05-juin-2028", "14-juil.-2028", "15-août-2028", "01-nov.-2028", "11-nov.-2028", "25-déc.-2028", "26-déc.-2028",

    # 2029
    "01-janv.-2029", "30-mars-2029", "02-avr.-2029", "01-mai-2029", "08-mai-2029", "10-mai-2029", "20-mai-2029",
    "21-mai-2029", "14-juil.-2029", "15-août-2029", "01-nov.-2029", "11-nov.-2029", "25-déc.-2029", "26-déc.-2029",

    # 2030
    "01-janv.-2030", "19-avr.-2030", "22-avr.-2030", "01-mai-2030", "08-mai-2030", "30-mai-2030", "09-juin-2030",
    "10-juin-2030", "14-juil.-2030", "15-août-2030", "01-nov.-2030", "11-nov.-2030", "25-déc.-2030", "26-déc.-2030"
  ]

# Convert holiday strings to datetime objects
HOLIDAY_DATES = [datetime.strptime(date, "%d-%b-%Y").date() for date in holidayDates]
