from setuptools import setup

APP = ['auto_email_mac.py']
DATA_FILES = [
    'attachment.pdf',
    'date.txt',
    '.env',
    'template.docx',
    'client_email.html',
    'facture_email.html',
    'notary_email.html',
    'signature.html',
]
OPTIONS = {
    'argv_emulation': False,  # Set to True if you want to emulate command line arguments for GUI apps
    'iconfile': 'bot.ico',  # macOS uses .icns format for icons, you may need to convert your .ico file
    'plist': {
        'CFBundleName': 'AutoEmail',
        'CFBundleDisplayName': 'AutoEmail',
        'CFBundleGetInfoString': "Your application description",
        'CFBundleVersion': "1.0.0",
        'CFBundleShortVersionString': "1.0.0",
    },
    'resources': DATA_FILES,  # Include additional data files with your app
}

setup(
    app=APP,
    name='AutoEmail',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
