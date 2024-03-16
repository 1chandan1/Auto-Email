from setuptools import setup

APP = ["auto_email.py"]
DATA_FILES = [
    "date.txt",
    ".env",
    ("static", ["static/"]),
]
OPTIONS = {
    "argv_emulation": True,
    'packages': ['charset_normalizer'],
    "iconfile": "static/bot.ico",  # macOS uses .icns format for icons, you may need to convert your .ico file
    "plist": {
        "CFBundleName": "AutoEmail",
        "CFBundleDisplayName": "AutoEmail",
        "CFBundleGetInfoString": "Your application description",
        "CFBundleVersion": "1.0.0",
        "CFBundleShortVersionString": "1.0.0",
    },
    "resources": DATA_FILES,  # Include additional data files with your app
}

setup(
    app=APP,
    name="AutoEmail",
    data_files=DATA_FILES,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"]
)
