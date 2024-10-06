from datetime import datetime
from src.utils import resource_path


SIGNATURE_TEMPLATE_PATH = resource_path("templates/signature.html")
NOTARY_EMAIL_TEMPLATE1_PATH = resource_path("templates/notary_email_1.html")
NOTARY_EMAIL_TEMPLATE2_PATH = resource_path("templates/notary_email_2.html")
UNDERTAKER_EMAIL_TEMPLATE1_PATH = resource_path("templates/undertaker_email_1.html")
UNDERTAKER_EMAIL_TEMPLATE2_PATH = resource_path("templates/undertaker_email_2.html")


INTERN_SHEET_KEY = "1cveFT3BvSJ9d-PvBdyGvZ7Fd_0XYDlNafOpTXUiz_eI"
INVOICE_SHEET_KEY = "1nuWrAZB4XF2Jlo_IaLPxRra-13H_fQffbTT1LBJyqVM"
DEATH_CERTIFICATES_SHEET_KEY = "1e4GzXCftJYFRbh3xKWnvRK8zE-FjOUut7FinhFj-2ug"
NOTARY_ANNUAIRE_SHEET_KEY = "1NBWDbmuXHKr6yWsEvxJhio4uaUPKol6_dJvtgKJCDhc"
UNDERTAKER_ANNUAIRE_SHEET_KEY = "12xP7d6R-lhoT39z2b4Jk2Ap07bP_nN6m1j2UXMHaVuk"
ANNUAIRE_WORKSHEET_NAME = "Notaire annuaire"
UNDERTAKER_WORKSHEET_NAME ="PF Annuaire"
SCHEDULED_WORKSHEET_NAME = "Scheduled email"


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
