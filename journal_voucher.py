import win32com.client
import os
import datetime as dt
import pandas as pd
import variables

def create_journal_voucher(wb, RU, year, period, methodology):
    """ Function to build up the reconciliation sheet structure"""

    sheet = wb.Worksheets.Add()

    sheet.Select()




    # sheet.Cells(8, 2).Value = f"Scale parameter is $ 1,000 (k) USD : $1 USD"
