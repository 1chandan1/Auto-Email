import subprocess
import platform
import pytesseract
from openai import OpenAI

from utils import *


openai_client = OpenAI(api_key = get_gpt_key())


def ask_date(image):
    text = pytesseract.image_to_string(image, lang="fra")
    prompt = (
        text
        + """

    the text is in french
    get date of birth as (DOB), date of death as (DOD) from it
    date in format "DD/MM/YYYY"
    if not found then ""
    json
    {
        "DOB":"DD/MM/YYYY",
        "DOD":"DD/MM/YYYY"
    }
    """
    )
    response = openai_client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                ],
            }
        ],
        response_format={"type": "json_object"},
        max_tokens=300,
    )
    try:
        result = eval(response.choices[0].message.content)
        return result
    except:
        return {}

def setup_tesseract():
    os_name = platform.system()
    if os_name == "Windows":
        if os.path.exists("C:/Program Files/Tesseract-OCR"):
            tesseract_path = "C:/Program Files/Tesseract-OCR/tesseract.exe"
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            if "fra" in pytesseract.get_languages():
                return
        else:
            pass
        print("!! tesseract is not installed !!")
        print(
            "Download and install tesseract : https://github.com/UB-Mannheim/tesseract/wiki"
        )
        print("Select French language during installation")
        input()
        sys.exit()
    else:
        try:
            result = subprocess.run(
                ["tesseract", "--version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode == 0:
                if "fra" in pytesseract.get_languages():
                    return
        except FileNotFoundError:
            pass
        print("!! tesseract-fra is not installed !!")
        print('Install tesseract-fra with this command : "brew install tesseract-fra"')
        input()
        sys.exit()


setup_tesseract()
