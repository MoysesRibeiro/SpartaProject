"""This is the file that needs to be updated in case variables need to be changed"""

import datetime as dt
import os
import base64
import pandas as pd
from tkinter import filedialog
from email_settings import GetEmail


__SCALLING_FACTOR__ = 1


def create_template_adjustment_table():
    create_folder()
    p = f'{os.environ["USERPROFILE"]}\\.SpartaTool'

    df1 = pd.DataFrame(columns=['ID', 'RU', 'YEAR', 'PERIOD', 'PERM_AMOUNT', 'DESCRIPTION'],
                       data=[[0, '0000', 2999, '09', 9999, 'Template default text']])
    df2 = pd.DataFrame(columns=['ID', 'RU', 'YEAR', 'PERIOD', 'TEMP_AMOUNT', 'DESCRIPTION'],
                       data=[[0, '0000', 2999, '09', 9999, 'Template default text']])
    df3 = pd.DataFrame(columns=['ID', 'RU', 'YEAR', 'PERIOD', 'CR_AMOUNT', 'DESCRIPTION'],
                       data=[[0, '0000', 2999, '09', 9999, 'Template default text']])

    writer = pd.ExcelWriter(f'{p}\\template.xlsx', engine="xlsxwriter", mode='w')

    df1.to_excel(writer, sheet_name="Perms", index=False)
    df2.to_excel(writer, sheet_name="Temps", index=False)
    df3.to_excel(writer, sheet_name="Credits", index=False)

    writer.close()


def create_folder():
    p = f'{os.environ["USERPROFILE"]}\\.SpartaTool'
    if not os.path.isdir(p):
        os.system(f'mkdir {p}')


def get_end_of_month_date(year, period):
    year = year if period < 12 else year + 1
    period = period + 1 if period < 12 else 1

    date = dt.datetime(year, period, 1)
    date = date - dt.timedelta(days=1)
    return date


def ask_for_path():
    create_folder()

    p = f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_adjustment_path.txt'
    if file_exists(p):
        b, t = is_file_not_empty(p)
        if b:
            if file_exists(t):
                return str(t)
            else:
                print("Warning Number 002: '", t, "' is not a valid file path, please inform new path:")
                new_path = filedialog.askopenfile(title = 'Select sparta master data path').name
                with open(f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_adjustment_path.txt', 'w') as f:
                    f.write(new_path)
                    f.close()


                ask_for_path()
        else:
            print(
                f"""Warning Number 001 : The log file on {os.environ["USERPROFILE"]}\\.SpartaTool\\
                    sparta_adjustment_path.txt, does not contain a valid file_path, please inform the
                    adjustment table .xlsx file to look for.""")
            text_to_input =  filedialog.askopenfile(title = 'Select sparta master data path').name
            with open(f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_adjustment_path.txt', 'w') as f:
                f.write(text_to_input)
                f.close()
            ask_for_path()
    else:
        set_text_file()
        ask_for_path()

def get_correct_email():
    app = GetEmail()
    app.mainloop()
    get_email()


def get_email():

    p = f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_email.txt'
    if file_exists(p):
        b, t = is_file_not_empty(p)
        if b:
            return t
        else:
            app = GetEmail()
            app.mainloop()

            get_email()
    else:
        set_email_file()
        get_email()


def is_file_not_empty(file):
    p = file
    with open(p, 'r') as f:
        string = f.read()

    return True if len(string) != 0 else False, string


def set_text_file():
    p = f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_adjustment_path.txt'
    if not file_exists(p):
        with open(p, 'w') as f:
            f.close()
    else:
        pass

def set_email_file():
    p = f'{os.environ["USERPROFILE"]}\\.SpartaTool\\sparta_email.txt'
    if not file_exists(p):
        with open(p, 'w') as f:
            f.close()
    else:
        pass

def file_exists(p):
    return os.path.isfile(p)


def get_path():
    return ask_for_path()

def generate_icon(icon: str) -> str:
    icon_data = base64.b64decode(icon)
    temp_file = os.environ["USERPROFILE"] + r"\.SpartaTool\icon.ico"
    icon_file = open(temp_file, "wb")
    icon_file.write(icon_data)
    icon_file.close()
    return temp_file

perms_path = ""
credits_path = perms_path
temps_path = perms_path

dictionary_main_vs_concept = {18: "Miscellaneous",
                              20: "Inventory",
                              21: "Inventory",
                              28: "Inventory",
                              41: "Pension",
                              45: "Lease ROUA",
                              49: "Depreciation",
                              58: "Miscellaneous",
                              64: "Miscellaneous",
                              71: "Lease Liability",
                              195:"Other",
                              205:"Inventory",
                              220:"Inventory",
                              302:"Depreciation",
                              308:"Depreciation",
                              306:"Depreciation",
                              366:"Depreciation",
                              390:"Lease",
                              395:"Lease",
                              598:"Pension",
                              770:"Pension",
                              788:"Other"
                              }

dictionary_base_vs_concept = {710010:"Pension",
                              710011:"Pension",
                              710012:"Pension",
                              710020:"Depreciation",
                              710035:"Carryforward",
                              710040:"Other",
                            465010:"Pension",
                            465011:"Pension",
                            465012:"Pension",
                            465020:"Depreciation",
                            465022:"Lease",
                            465023:"Lease",
                            465030:"Carryforward",
                            465050:"Other",
                            465200:"Inventory"

                              }
