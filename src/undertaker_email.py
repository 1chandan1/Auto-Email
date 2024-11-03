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


def undertaker_email(user: GoogleServices):
    user.print_details()
    print_center("Undertaker Email")
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
    print_center("Undertaker Email")
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
            send_undertaker_emails(user, spreadsheet)
            user.print_details()
            print_center("Undertaker Email")
            print()
            print("\n\nTask Completed")
            countdown("Exit", 5)
            sys.exit()
        elif choice == "2":
            undertaker_email(user)
            break
        elif choice == "q":
            break
    sleep(0.1)


def send_undertaker_emails(user: GoogleServices, spreadsheet: gspread.Spreadsheet):
    undertaker_annuraie_sheet = user.gc.open_by_key(UNDERTAKER_ANNUAIRE_SHEET_KEY)
    death_certificates_sheet = user.gc.open_by_key(DEATH_CERTIFICATES_SHEET_KEY)
    death_certificates_worksheet = death_certificates_sheet.get_worksheet_by_id(0)
    all_image_data = death_certificates_worksheet.get_all_values()
    undertaker_annuraie_worksheet = undertaker_annuraie_sheet.worksheet(
        UNDERTAKER_WORKSHEET_NAME
    )
    undertaker_scheduling_worksheet = undertaker_annuraie_sheet.worksheet(
        SCHEDULED_WORKSHEET_NAME
    )
    worksheet = spreadsheet.get_worksheet(0)
    sheet_data = worksheet.get_values()
    header = sheet_data[0]
    rows = sheet_data[1:]
    df = pd.DataFrame(rows, columns=header)
    df.index += 2
    status_col_index = df.columns.get_loc("Status") + 1
    comment_column_index = status_col_index + 2
    all_annuraie_data = undertaker_annuraie_worksheet.get_all_values()
    all_scheduled_data = undertaker_scheduling_worksheet.get_all_values()
    all_cases = get_all_cases(all_annuraie_data, all_scheduled_data)
    for index, row in df.iterrows():
        try:
            if row["Status"] == "à envoyer":
                undertaker_email: str = row["Email"].split("\n")[0].strip()
                person_full_name: str = row["Name"].strip()
                _, person_last_name = get_fname_lname(person_full_name)
                declarant_name = row["Declarant Name"].strip()
                person_dod: str = row["Date Of Death"].strip()
                if not all(
                    [
                        person_last_name.strip(),
                        person_dod,
                        person_dod != "N/A",
                        is_valid_email(undertaker_email),
                    ]
                ):
                    worksheet.update_cell(index, comment_column_index, "Problem")
                    continue

                undertaker_sheet_index = find_row_by_name(
                    all_annuraie_data, declarant_name
                )

                email_row_indices = find_rows_by_email(
                    all_annuraie_data, undertaker_email
                )
                
                undertaker_status = get_status(all_annuraie_data, email_row_indices)
                if not undertaker_sheet_index:
                    entreprise = find_entreprise_by_email(
                        all_annuraie_data, undertaker_email
                    )
                    new_row = [
                        entreprise,
                        None,
                        declarant_name,
                        row.get("City"),
                        row.get("Street"),
                        row.get("Phone"),
                        undertaker_email,
                        undertaker_status,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        user.email,
                    ]
                    annuraie_updated_row = undertaker_annuraie_worksheet.append_row(
                        new_row, value_input_option="USER_ENTERED", table_range="A:A"
                    )
                    updated_range = annuraie_updated_row["updates"]["updatedRange"]
                    undertaker_sheet_index = extract_row_number(updated_range)
                    all_annuraie_data.append(new_row)
                    email_row_indices.append(undertaker_sheet_index)
                    worksheet.update_cell(
                        index, comment_column_index, "New Declarant added"
                    )

                annuraie_sheet_row = all_annuraie_data[undertaker_sheet_index - 1]
                
                user.print_details()
                print_center("Undertaker Email")
                print()
                print_center(f"Google Sheet : {spreadsheet.title}")
                print()
                print_center("Sending All Emails\n\n")
                print(f"Annuraie sheet Row    :    {undertaker_sheet_index}") 
                print(f"Target Sheet Row  :    {index}\n")
                print(f"Person Name       :    {person_full_name}")
                print(f"Person Last Name  :    {person_last_name}\n")
                print(f"DON               :    {person_dod}")
                print(f"To                :    {undertaker_email}\n")

                name_for_checking = re.sub(
                    r"[ ,\-\n]", "", unidecode(person_full_name).lower().strip()
                )

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
                    if undertaker_status == "Not contacted":
                        countdown("Sending Email in", random.randint(120, 180))
                        print("\nSending Email...")
                        status = None
                        for _ in range(3):
                            message = create_undertaker_message(
                                user,
                                undertaker_email,
                                person_full_name,
                                person_last_name,
                                person_dod,
                                image_data,
                                image_name,
                            )
                            status = user.send_email(message)
                            if status:
                                break
                            countdown("Trying to send email again", 5)

                        if status:
                            worksheet.update_cell(index, status_col_index, "envoyé")
                            today_date = datetime.now().date().strftime("%d-%b-%Y")
                            undertaker_annuraie_worksheet.update_acell(
                                f"I{undertaker_sheet_index}", today_date
                            )
                            undertaker_annuraie_worksheet.update_acell(
                                f"L{undertaker_sheet_index}", person_full_name
                            )
                            all_annuraie_data[undertaker_sheet_index - 1][8] = today_date
                            all_annuraie_data[undertaker_sheet_index - 1][11] = person_full_name
                            for i in email_row_indices:
                                undertaker_annuraie_worksheet.update_acell(
                                    f"H{i}",
                                    "Contacted / pending answer",
                                )
                                all_annuraie_data[i - 1][7] = "Contacted / pending answer"
                            print("\nSuccess")
                        else:
                            worksheet.update_cell(index, comment_column_index, f"Error")
                    elif undertaker_status in (
                        "Contacted / pending answer",
                        "Cooperating",
                        "Collègue coopérant"
                    ):

                        previous_scheduled_date = annuraie_sheet_row[8]
                        new_date_text = None
                        if undertaker_status == "Cooperating" or undertaker_status == "Collègue coopérant":
                            # List to store all previous scheduled dates
                            previous_dates = []
                            try:
                                previous_dates.append(
                                    datetime.strptime(
                                        previous_scheduled_date, "%d-%b-%Y"
                                    ).date()
                                )
                            except:
                                pass

                            for scheduled_data in all_scheduled_data:
                                if undertaker_email in scheduled_data:
                                    try:
                                        # Extract the date and append to the list
                                        previous_scheduled_date = datetime.strptime(
                                            scheduled_data[8], "%d-%b-%Y"
                                        ).date()
                                        previous_dates.append(previous_scheduled_date)
                                    except:
                                        # Handle any issues with date parsing, skip if invalid
                                        continue
                            today = datetime.now().date()
                            # Get the maximum date from the list of previous dates if available
                            if previous_dates:
                                max_previous_date = max(previous_dates)
                                new_date = max_previous_date + relativedelta(months=+2)
                                if new_date < today:
                                    new_date = today + relativedelta(months=+2)
                            else:
                                # Fallback if no previous dates are found
                                new_date = today + relativedelta(months=+2)

                            # Check if new_date falls on a weekend or holiday, and adjust accordingly
                            while (
                                new_date.weekday() in (5, 6)
                                or new_date in HOLIDAY_DATES
                            ):
                                if new_date.weekday() == 5:  # Saturday
                                    new_date += relativedelta(days=2)
                                elif new_date.weekday() == 6:  # Sunday
                                    new_date += relativedelta(days=1)
                                elif new_date in HOLIDAY_DATES:  # Holiday
                                    new_date += relativedelta(days=1)

                            # Convert new_date to the required format
                            new_date_text = new_date.strftime("%d-%b-%Y")

                        new_schedule_row = [
                            None,
                            declarant_name,
                            None,
                            "Scheduled" if new_date_text else None,
                            person_full_name,
                            person_dod,
                            user.email,
                            undertaker_email,
                            new_date_text,
                            None,
                            None,
                            None,
                            image_link,
                        ]
                        undertaker_scheduling_worksheet.append_row(
                            new_schedule_row,
                            value_input_option="USER_ENTERED",
                            table_range="A:A",
                        )
                        all_scheduled_data.append(new_schedule_row)
                        worksheet.update_cell(
                            index,
                            comment_column_index,
                            f"Scheduled {new_date_text or ''}",
                        )
                        worksheet.update_cell(index, status_col_index, "draft")
                        print(f"Scheduled {new_date_text}")
                    elif undertaker_status == "Not cooperating":
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
        re.sub(r"[ ,\-\n]", "", unidecode(row[11]).lower().strip())
        for row in all_annuraie_data
    ]
    all_cases = list(set(all_annuraie_cases + all_scheduled_cases))
    return all_cases


def find_row_by_name(annuraie_data, declarant_name):
    pattern = r"[ ,\-\n]"
    declarant_name = unidecode(re.sub(pattern, "", declarant_name)).lower()
    for index, row in enumerate(annuraie_data, start=1):
        if declarant_name == unidecode(re.sub(pattern, "", row[2])).lower():
            return index
    return None


def find_entreprise_by_email(annuraie_data, email):
    for row in annuraie_data:
        if email in row:
            return row[0]
    return None


def get_status(data, row_indices):
    if row_indices:
        all_status = [data[index - 1][7] for index in row_indices]
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


def create_undertaker_message(
    user: GoogleServices,
    to: str,
    person_full_name: str,
    person_last_name: str,
    person_dod: str,
    attachment: io.BytesIO = None,  # attachment is io.BytesIO
    attachment_filename: str = None,  # Provide the filename separately
):
    message = MIMEMultipart()
    message["From"] = f"Klero Genealogie <{user.email}>"
    message["To"] = to
    message["Subject"] = f"Obsèque {person_last_name} - Heritage non transmis - Demande de mise en relation avec la famille"
    

    if attachment and attachment_filename:
        template_path = UNDERTAKER_EMAIL_TEMPLATE2_PATH
    else:
        template_path = UNDERTAKER_EMAIL_TEMPLATE1_PATH

    with open(template_path, "r", encoding="utf-8") as file:
        html_template = file.read()
        message_html = html_template.format(
            person_full_name=person_full_name, person_dod=person_dod
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
