import snowflake.connector as sf
import sys
import threading
import time
from tkinter import IntVar
import datetime as dt
import variables
'''
Basic Structure
'''



class SnowSQL:
    def __init__(self, user):
        self.condition = IntVar()
        self.condition.set(0)
        try:
            self.con = sf.connect(
                user=user,
                account='xom-fin',  # account identifier
                authenticator='externalbrowser',
                warehouse='BI_FED_WH'  # warehouse can be changed according to database requiremets
            )

        except Exception as e:
            print('FED connection error... due to ', e)
            variables.get_correct_email()
            sys.exit()

    def execute_query(self, query, ru, year, period):
        self.condition = False
        start = dt.datetime.now()
        print(f"\rQuerying data for {ru} on {period}-{year}")
        th = threading.Thread(target=loading_bar, args=(self,))
        th.start()
        try:
            cursor = self.con.execute_string(query)[0]
            print(f"\rQuery for {ru} on {period}-{year} successfull")
            self.condition = True
            return cursor
        except Exception as e:
            print(f"\rQuery for {ru} on {period}-{year} failed due to ", e)
            return []


def loading_bar(condition):
    x = 0
    con = condition.condition
    loadings = ["Loading.", "Loading..", "Loading..."]
    while not con:
        x = 0 if x > 2 else x
        sys.stdout.write(f'\r{loadings[x]}')
        time.sleep(1)
        x += 1
        con = condition.condition
