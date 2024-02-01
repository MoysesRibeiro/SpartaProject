"""
Created on Fri Aug 18th, 2023 16:05:35
Updated on Thu Jan 31st, 2024 15:19:30
Name = "SpartaProject"
Version = "1.0.0"
Description = "To calculate tax provision"
Authors = MoysesRibeiro & FelipeKnob
License = "ExxonMobil Controllers-Tax"

This is the version on moyses computer.

"""


from main_gui import MainGUI
from email_settings import GetEmail
import variables
import os

if __name__ == '__main__':
    os.system("COLOR 2F")

    email = variables.get_email()

    print(email)


    app = MainGUI(email)
    app.resizable(True, True)
    app.mainloop()

"""
to build the .exe file use code to follow:
pyinstaller --additional-hooks-dir . --noconfirm  --onefile --icon=spartan3.ico main.py
see documentation at: https://github.com/snowflakedb/snowflake-connector-python/issues/698

Packages and respective versions used in this project:
altgraph==0.17.3
appdirs==1.4.4
asn1crypto==1.5.1
beautifulsoup4==4.12.2
certifi==2023.7.22
cffi==1.15.1
charset-normalizer==2.1.1
cryptography==38.0.4
et-xmlfile==1.1.0
filelock==3.12.2
frozendict==2.3.8
html5lib==1.1
idna==3.4
lxml==4.9.3
multitasking==0.0.11
numpy==1.25.2
openpyxl==3.1.2
oscrypto==1.3.0
packaging==23.1
pandas==1.5.3
pefile==2023.2.7
platformdirs==3.8.1
pyarrow==8.0.0
pycparser==2.21
pycryptodomex==3.18.0
pyinstaller==5.13.0
pyinstaller-hooks-contrib==2023.7
PyJWT==2.8.0
pyOpenSSL==22.1.0
pypiwin32==223
pyspnego==0.10.2
python-dateutil==2.8.2
pytz==2023.3
pywin32==306
pywin32-ctypes==0.2.2
requests==2.31.0
requests-ntlm==1.2.0
requests-toolbelt==1.0.0
six==1.16.0
snowflake-connector-python==2.8.3
sortedcontainers==2.4.0
soupsieve==2.4.1
sspilib==0.1.0
tomlkit==0.12.1
typing_extensions==4.7.1
tzdata==2023.3
urllib3==1.26.16
webencodings==0.5.1
XlsxWriter==3.1.2
yfinance==0.2.28
"""
