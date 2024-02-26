import win32com.client
import os
import datetime as dt
import pandas as pd
import variables


def open_excel_file(file_name):
    """Function to create/instantiate excel file in background"""

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = 0
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(file_name)
    xl.ActiveWindow.DisplayGridlines = False

    """Write into excel HEADER"""
    # sheet = wb.Worksheets.Add()
    # wb.Worksheets(1).Name = 'Cover'

    # return xl, wb, sheet

    return xl, wb


def create_cover(wb, RU, year, period, methodology):
    """ Function to build up the reconciliation sheet structure"""

    sheet = wb.Worksheets.Add()

    sheet.Select()
    sheet.Cells(2, 2).Value = "Tax Analysis Tool"
    sheet.Cells(3, 2).Value = "This working paper was generated by a tool, designed to assist in tax analysis"
    sheet.Cells(4,
                2).Value = f"""Created by : {os.environ['USERNAME'].upper()} at :{
    dt.datetime.strftime(dt.datetime.now(), '%B %dth, %Y  %H:%M:%S')} """

    sheet.Cells(6, 2).Value = "Parameters used in this tool:"
    sheet.Cells(7, 2).Value = f"RU : {RU}  Year : {year}  Period: {period}"
    sheet.Cells(8, 2).Value = f"Methodology used in this working paper : {methodology}"
    sheet.Cells(9,
                2).Value = f"Which means that the IBIT used to calculate tax is considering the {methodology} trial balance"

    # sheet.Cells(8, 2).Value = f"Scale parameter is $ 1,000 (k) USD : $1 USD"


def create_total_tax_mask(xl, wb, dataframe_dictionary: dict,
                          tax_rate: float, methodology: str, country: str) -> None:
    """
    Creates a simple form starting from net income until the total tax provision.

    :param xl: Excel application pipe through pywin32
    :param wb: xl.Workbook -> a workbook object from Excel
    :param dataframe_dictionary: a dictionary containing several dataframes
    :param tax_rate: the tax rate from the respective country
    :param methodology: retrieving data from local income statement or GAAP income statement
    :param country: country in which the company operates (makes a difference only if it is "US")

    :returns: Nothing, only a void function.
    
    """

    # sheet = wb.Worksheets.Add(After=wb.Worksheets(1))
    # wb.Worksheets(2).Name = 'A.TT'

    tb = dataframe_dictionary.get('trial_balance_local')
    tb_GAAP = dataframe_dictionary.get('trial_balance_GAAP')

    if methodology == "Local":
        net_income = -tb_GAAP['AMOUNT'].sum() if country == 'US' else -tb['AMOUNT'].sum()

        tax = +tb[tb['G/L Acct'].str[:2] == '89']['AMOUNT'].sum() + \
              tb[tb['G/L Acct'].str[:2] == '90']['AMOUNT'].sum() + \
              tb[tb['G/L Acct'].str[:4] == 'N089']['AMOUNT'].sum() + \
              tb[tb['G/L Acct'].str[:4] == 'N090']['AMOUNT'].sum()

    else:
        net_income = -tb_GAAP['LC_AMOUNT'].sum()
        tax = +tb_GAAP[tb_GAAP['G/L Acct'].str[:2] == '89']['LC_AMOUNT'].sum() + \
              tb_GAAP[tb_GAAP['G/L Acct'].str[:2] == '90']['LC_AMOUNT'].sum()

    sheet = wb.Worksheets.Add(Before=wb.Worksheets(1))
    sheet.Select()

    # HEADERS
    sheet.Cells(4, 2).Value = "Working paper : 1.Total Tax"
    sheet.Cells(5, 2).Value = "Subject: Total Tax calculation"
    sheet.Cells(4, 11).Value = f"Prepared by : {os.environ['USERNAME'].upper()}"
    sheet.Cells(5, 11).Value = f"Date : {dt.datetime.strftime(dt.datetime.now(), '%B %dth, %Y  %H:%M:%S')}"

    sheet.Cells(8,
                2).Value = "The income tax calculation takes into account IBIT + permanent adjustments materially " \
                           "over $ 1M USD and applies the income tax statutory rate over it "

    # SECTION 1 : TOTAL TAX IN LOCAL CURRENCY
    sheet.Cells(10, 2).Value = "SECTION 1 : TOTAL TAX IN LOCAL CURRENCY"
    sheet.Cells(10, 2).Font.Bold = True
    sheet.Cells(14, 2).Value = "1.Net Income Help/(Hurt)"
    sheet.Cells(12, 5).Value = "Total"

    sheet.Cells(14, 5).Value = net_income
    sheet.Cells(16, 5).Value = tax

    sheet.Cells(16, 3).Value = "' + Taxes"

    sheet.Cells(18, 2).Value = "IBIT"
    sheet.Cells(18, 5).Value = net_income + tax

    # FSI SECTION
    sheet.Cells(20, 3).Value = "FSI"
    sheet.Cells(21, 3).Value = "Non-Taxable income amount"
    sheet.Cells(21, 5).Value = 0
    sheet.Cells(22, 3).Value = "Taxable Income %"
    sheet.Cells(22, 5).Value = 1
    sheet.Cells(22, 5).NumberFormat = "0.00%"
    sheet.Cells(22, 5).Interior.Color = 16776960
    sheet.Cells(21, 5).Interior.Color = 16776960

    sheet.Cells(24, 2).Value = "IBIT after FSI"
    sheet.Cells(24, 5).Value = "=+(E18-E21)*E22"

    sheet.Cells(26, 2).Value = "(+) 2. Permanent Differences Add/(Subtract)"
    sheet.Cells(26, 5).Value = dataframe_dictionary.get('perms_from_excel')[
                                   'PERM_AMOUNT'].sum() / variables.__SCALLING_FACTOR__
    sheet.Cells(28, 2).Value = "(=) Taxable Income"
    sheet.Cells(28, 5).Value = "=+E26+E24"
    sheet.Cells(30, 3).Value = "(X) Statutory Tax rate"
    sheet.Cells(30, 5).Value = tax_rate / 100
    sheet.Cells(30, 5).Style = "Percent"
    sheet.Cells(30, 5).NumberFormat = "0.0%"
    sheet.Cells(32, 2).Value = "(=) Tax before Credits"
    sheet.Cells(32, 5).Value = "=E28*E30"
    sheet.Cells(34, 2).Value = "(-) Tax Credits"
    sheet.Cells(34, 5).Value = dataframe_dictionary.get('credits_from_excel')[
                                   'CR_AMOUNT'].sum() / variables.__SCALLING_FACTOR__
    sheet.Cells(37, 2).Value = "(=) Total Tax"
    sheet.Cells(37, 5).Value = "=E32-E34"

    # SECTION 2 : GAAP DEFERRED TAX
    sheet.Cells(40, 2).Value = "SECTION 2 : GAAP DEFERRED TAX"
    sheet.Cells(40, 2).Font.Bold = True

    sheet.Cells(42, 2).Value = "Current Tax in USD"
    if methodology == "Local":
        sheet.Cells(42, 5).Value = "=+E37/'Exchange Rate'!B4-E43"

    elif methodology == "GAAP":
        sheet.Cells(42, 5).Value = "+E37/'Exchange Rate'!B4-E43-E44"

    sheet.Cells(43, 2).Value = "Deferred Tax USD"
    sheet.Cells(43, 5).Value = "=-IFERROR(SUM(Temps!F:F)/'Exchange Rate'!B4*A.TT!E30,0)"

    sheet.Cells(44, 2).Value = "'+ EM Only Deferred Tax"
    sheet.Cells(44, 5).Value = dataframe_dictionary.get('trial_balance_delta_between_GAAPs')['TAX'].sum()

    sheet.Cells(46, 2).Value = "'= TOTAL TAX USD"
    sheet.Cells(46, 5).Value = "=SUM(E42:E44)"

    # SECTION 3 : FEDERAL VS. STATE
    sheet.Cells(48, 2).Value = "SECTION 3 : Tax Break-down"
    sheet.Cells(48, 2).Font.Bold = True

    sheet.Cells(51, 2).Value = "State Tax"
    sheet.Cells(51, 4).Value = 0.033 if country == 'US' else 0
    sheet.Cells(51, 4).NumberFormat = "0.0%"
    sheet.Cells(51, 5).Value = "=+(E24*D51)/'Exchange Rate'!B4"

    sheet.Cells(52, 2).Value = "Federal Tax"
    sheet.Cells(52, 5).Value = "=+E46-E51"

    sheet.Cells(54, 2).Value = "'+ TOTAL TAX USD"
    sheet.Cells(54, 5).Value = '=+E51+E52'

    sheet.Cells(56, 2).Value = "SECTION 4 : Trend Analysis"
    sheet.Cells(56, 2).Font.Bold = True
    sheet.Cells(57,
                2).Value = "Disclaimer: this is not required by audits or internal controls. It's purpose is limited only to assist in tax trends analysis with visuals."
    sheet.Cells(57, 2).Font.Size = 9

    sheet.Cells(97, 2).Value = "Commentary:"
    # FORMATING
    sheet.Columns("A:A").ColumnWidth = 0.94
    sheet.Columns("B:B").ColumnWidth = 0.63
    sheet.Columns("D:D").ColumnWidth = 40.22
    sheet.Columns("H:H").ColumnWidth = 25.55
    xl.ActiveWindow.DisplayGridlines = False
    sheet.Columns("E:E").NumberFormat = "#,##0"

    if country != "US":
        pass


def create_cf216_mask(xl, wb, dataframe_dictionary: dict, currency: str) -> None:
    """
    Creates a simulated CF216 mask with data in it to upload to DFX.

    :param xl: Excel application pipe through pywin32
    :param wb: xl.Workbook -> a workbook object from Excel
    :param dataframe_dictionary: a dictionary containing several dataframes
    :param currency: the company local currency
    :returns: Nothing, only a void function

    """

    sheet = wb.Worksheets.Add(Before=wb.Worksheets(4))
    # wb.Worksheets(4).Name = 'C.CF216'
    # sheet = wb.Worksheets.Add()
    sheet.Select()
    tb = dataframe_dictionary.get('trial_balance_GAAP')

    sheet.Cells(1, 1).Value = "RUNDATE:08/24/23 RUNTIME:0803 JOBNAME:BFX86EXD SUBPER:23BXM06 CG:(01 "
    sheet.Cells(2, 1).Value = "This spreadsheet is Comment Enabled"
    sheet.Cells(3, 1).Value = "RU:1556 ExxonMobil Mexico, S.A. DE C.V."
    sheet.Cells(4, 1).Value = "FORM:0216 EFFECTIVE TAX RATE RECONCILIATION"

    sheet.Cells(5, 1).Value = "SEG:0100 Corporate Total"
    sheet.Cells(6, 1).Value = "|--------------------------------------------|"
    sheet.Cells(6, 2).Value = "'===COL 01==="
    sheet.Cells(6, 3).Value = "'===COL 02==="
    sheet.Cells(6, 4).Value = "'===COL 03==="
    sheet.Cells(6, 5).Value = "'===COL 04==="

    sheet.Cells(7, 1).Value = "|--------------------------------------------|"
    sheet.Cells(7, 2).Value = "INC TAX US"
    sheet.Cells(7, 3).Value = "IBIT US"
    sheet.Cells(7, 4).Value = "IT NONUS"
    sheet.Cells(7, 5).Value = "IBIT NONUS"

    sheet.Cells(10, 1).Value = "L0010 Total Inc Tax/IBIT incl NCI & Adj"
    sheet.Cells(10, 2).Value = 0
    sheet.Cells(10, 3).Value = 0
    sheet.Cells(10, 4).Value = "=+A.TT!E46"
    sheet.Cells(10, 5).Value = -tb['AMOUNT'].sum() + tb[tb['G/L Acct'].str[:2] == '89']['AMOUNT'].sum() + \
                               tb[tb['G/L Acct'].str[:2] == '90']['AMOUNT'].sum()

    sheet.Cells(11, 1).Value = "'  +++ Reconciling Factors                 +++"

    sheet.Cells(12, 1).Value = "L0020 Tax Rate and Law Changes"
    sheet.Cells(12, 2).Value = 0
    sheet.Cells(12, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(12, 4).Value = 0
    sheet.Cells(12, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(13, 1).Value = "L0030 Prior Period Adjustments"
    sheet.Cells(13, 2).Value = 0
    sheet.Cells(13, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(13, 4).Value = 0
    sheet.Cells(13, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(14, 1).Value = "L0040 Withholding Taxes (DWT, IWT)"
    sheet.Cells(14, 2).Value = 0
    sheet.Cells(14, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(14, 4).Value = 0
    sheet.Cells(14, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(15, 1).Value = "L0060 Valuation Allowance"
    sheet.Cells(15, 2).Value = 0
    sheet.Cells(15, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(15, 4).Value = 0
    sheet.Cells(15, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(16, 1).Value = "L0070 Intercompany Tax Allocations  *"
    sheet.Cells(16, 2).Value = 0
    sheet.Cells(16, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(16, 4).Value = 0
    sheet.Cells(16, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(17, 1).Value = "L0080 Multiple Tax Rates            **"
    sheet.Cells(17, 2).Value = "XXXXXXXXXXX"
    sheet.Cells(17, 3).Value = 0
    sheet.Cells(17, 4).Value = "XXXXXXXXXXX"
    sheet.Cells(17, 5).Value = 0

    sheet.Cells(18, 1).Value = "L0090 Manufacturing Income Deduction"
    sheet.Cells(18, 2).Value = "XXXXXXXXXXX"
    sheet.Cells(18, 3).Value = 0
    sheet.Cells(18, 4).Value = "XXXXXXXXXXX"
    sheet.Cells(18, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(19, 1).Value = "L0100 Asset Mgmt Income @ Diff Rate **"
    sheet.Cells(19, 2).Value = 0
    sheet.Cells(19, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(19, 4).Value = 0
    sheet.Cells(19, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(20, 1).Value = "L0110 Reserved for Future Use"
    sheet.Cells(20, 2).Value = "XXXXXXXXXXX"
    sheet.Cells(20, 3).Value = "XXXXXXXXXXX"
    sheet.Cells(20, 4).Value = "XXXXXXXXXXX"
    sheet.Cells(20, 5).Value = "XXXXXXXXXXX"

    sheet.Cells(21, 1).Value = "L0120 Other Nondeductible/Perm Diff **"
    sheet.Cells(21, 2).Value = "XXXXXXXXXXX"
    sheet.Cells(21, 3).Value = 0
    sheet.Cells(21, 4).Value = "XXXXXXXXXXX"
    sheet.Cells(21, 5).Value = (dataframe_dictionary.get("perms_from_excel")['PERM_AMOUNT'].sum() /
                                dataframe_dictionary.get("exchange_rate")[
                                    'Close']) / variables.__SCALLING_FACTOR__ if currency != "USD" else \
        dataframe_dictionary.get("perms_from_excel")['PERM_AMOUNT'].sum() / variables.__SCALLING_FACTOR__

    sheet.Cells(22, 1).Value = "L0130 Equity Co Inc Tax & Income Adj"
    sheet.Cells(22, 2).Value = "XXXXXXXXXXX"
    sheet.Cells(22, 3).Value = 0
    sheet.Cells(22, 4).Value = "XXXXXXXXXXX"
    sheet.Cells(22, 5).Value = 0

    sheet.Cells(23, 1).Value = "L0140 Foreign Exchange"
    sheet.Cells(23, 2).Value = 0
    sheet.Cells(23, 3).Value = 0
    sheet.Cells(23, 4).Value = 0
    sheet.Cells(23, 5).Value = "=+A.TT!E18/'Exchange Rate'!B4-E10-SUM(B.GTDs!E:E)"

    sheet.Cells(23, 6).Value = "C"
    sheet.Cells(23, 6).Font.Bold = True
    sheet.Cells(23, 6).Font.Italic = True
    sheet.Cells(23, 6).Font.Color = -16776961

    sheet.Cells(24, 1).Value = "L0150 Other                         **"
    sheet.Cells(24, 2).Value = 0
    sheet.Cells(24, 3).Value = 0
    sheet.Cells(24, 4).Value = "=SUM(D10:D22)-E25*A.TT!E30"
    sheet.Cells(24, 5).Value = 0

    sheet.Cells(25, 1).Value = "L0200 After Reconcil - Sum L10 - L150"
    sheet.Cells(25, 2).Value = 0
    sheet.Cells(25, 3).Value = 0
    sheet.Cells(25, 4).Value = "=SUM(D10:D24)"
    sheet.Cells(25, 5).Value = "=SUM(E10:E24)"

    # FORMATING
    sheet.Columns("A:A").ColumnWidth = 55.52
    sheet.Columns("B:E").ColumnWidth = 12.0
    xl.ActiveWindow.DisplayGridlines = False
    sheet.Columns("B:E").NumberFormat = "#,##0"


def pandas_excel(file_name: str, dataframe_dictionary: dict) -> None:
    """
    Saves all dataframes in the dictionary to a specific excel file.

    :param file_name: name of the file to be saved
    :param dataframe_dictionary: dictionary of dataframes to be saved
    :return: Nothing, only a void function
    """

    writer = pd.ExcelWriter(os.environ["USERPROFILE"] + f"\\Desktop\\Sparta_Output\\{file_name}", engine="xlsxwriter")

    dataframe_dictionary.get("trial_balance_delta_between_GAAPs").to_excel(writer, sheet_name="B.GTDs")
    dataframe_dictionary.get("tidy_segmentation").to_excel(writer, sheet_name="D.SEGMENTATION_MASK")
    dataframe_dictionary.get("tidy_segmentation").to_excel(writer, sheet_name="D1.SEGMENTED_IBIT")
    dataframe_dictionary.get("trial_balance_GAAP").to_excel(writer, sheet_name="TB_GAAP")
    dataframe_dictionary.get("trial_balance_local").to_excel(writer, sheet_name="TB_LOCAL")
    dataframe_dictionary.get("perms_from_excel").to_excel(writer, sheet_name="Perms")
    dataframe_dictionary.get("temps_from_excel").to_excel(writer, sheet_name="Temps")
    dataframe_dictionary.get("credits_from_excel").to_excel(writer, sheet_name="Credits")
    dataframe_dictionary.get("GTD_detail").to_excel(writer, sheet_name="GTDs Detail")
    dataframe_dictionary.get("exchange_rate").to_excel(writer, sheet_name="Exchange Rate")
    # dataframe_dictionary.get("segmentation").to_excel(writer, sheet_name="SEGMENTED_IBIT_ORIGINAL")

    writer.close()


def create_memo_forex(wb) -> None:
    """
    Creates a memo on how forex for CF216 was calculated.
    It basically calculates it in a different way, in order
    to match the amount already calculated in CF216, to validate it.

    :param wb: workbook object from Excel
    :return: Nothing, void function
    """
    # sheet = wb.Worksheets.Add(After=xl.ActiveWorkbook.Worksheets(xl.ActiveWorkbook.Worksheets.Count))
    sheet = wb.Worksheets.Add(Before=wb.Worksheets(5))
    sheet.Select()
    sheet.Name = "C1.MEMO FOREX"

    # HEADERS
    sheet.Cells(1, 1).Value = "MEMO: Break-down of Forex Line on CF216"
    sheet.Cells(1, 1).Font.Bold = True

    sheet.Cells(2, 1).Value = "Add Forex Local"
    sheet.Cells(2,
                2).Value = '=-(SUMIF(TB_LOCAL!E:E,"95",TB_LOCAL!D:D)+' \
                           'SUMIF(TB_LOCAL!E:E,"96",TB_LOCAL!D:D))/' + "'Exchange Rate'" + '!B4'

    sheet.Cells(3, 1).Value = "Less Forex USD"
    sheet.Cells(3, 2).Value = '=-VLOOKUP("21959799",TB_GAAP!B:C,2,FALSE)'

    sheet.Cells(4, 1).Value = "Add Exchange Rate Delta"
    sheet.Cells(4,
                2).Value = """=-((C.CF216!E10-'C1.MEMO FOREX'!B3)-(A.TT!E18/'Exchange Rate'!B4
                -'C1.MEMO FOREX'!B2)-SUM(B.GTDs!E:E))"""

    sheet.Cells(6, 1).Value = "Total"
    sheet.Cells(6, 2).Value = "=+B2-B3+B4"

    sheet.Cells(2, 3).Value = " Local forex accounts 95 and 96 in local currency divided by exchange rate"
    sheet.Cells(3, 3).Value = " USD Forex = account 21959799 in USD"
    sheet.Cells(4,
                3).Value = " (IBIT GAAP less Forex USD) less " \
                           "(IBIT Local less Forex Local at closing rate) less GTDs base"
    sheet.Cells(6, 3).Value = "C"

    sheet.Columns("C:C").Font.Color = -5112692
    sheet.Cells(6, 3).Font.Color = -16776961
    sheet.Cells(6, 3).Font.Bold = True
    sheet.Cells(6, 3).Font.Italic = True


def format_cells(xl, wb):
    """ Format excel cells."""

    print("Formating excel file")

    sheets = wb.Sheets
    for i in sheets:
        i.Select()
        xl.ActiveWindow.DisplayGridlines = False
        # i.Columns("A:O").Font.Name = "EMprint"
        if i.Name == "A.TT":
            i.Cells(22, 5).NumberFormat = "0.00%"
            i.Cells(30, 5).NumberFormat = "0.00%"
            # to recalculate formula
            i.Cells(43, 5).Value = "=-IFERROR(SUM(Temps!F:F)/'Exchange Rate'!B4*A.TT!E30,0)"
            continue
        if i.Name == "C.CF216":
            i.Columns("A:O").Font.Name = "Courier New"
            continue
        if i.Name == "B.GTDs":
            i.Columns("A:E").ColumnWidth = 12
            continue

        if i.Name == "D.SEGMENTATION_MASK":
            i.Columns("A:M").ColumnWidth = 11
            # i.Range("A2:M7").NumberFormat = "#,##0"
            continue
        i.Columns("A:O").AutoFit()

    wb.Sheets("Cover").Select()


def format_tab_color(wb):
    wb.Sheets("A.TT").Tab.Color = 3434496
    wb.Sheets("B.GTDs").Tab.Color = 3434496
    wb.Sheets("C.CF216").Tab.Color = 3434496
    wb.Sheets("C1.MEMO FOREX").Tab.Color = 6723891
    wb.Sheets("D.SEGMENTATION_MASK").Tab.Color = 3434496
    wb.Sheets("D1.SEGMENTED_IBIT").Tab.Color = 6723891


def transform_in_us(wb):
    """ If RU is US, makes some sort of changes in CF 216"""
    wb.Sheets("B.GTDs").Delete()
    wb.Sheets("C1.MEMO FOREX").Delete()
    wb.Sheets("A.TT").Cells(44, 5).Value = 0
    wb.Sheets("A.TT").Cells(43, 5).Value = "=-IFERROR(SUM(Temps!F:F)/'Exchange Rate'!B4*21%,0)"
    wb.Sheets("C.CF216").Cells(23, 5).Value = "=+A.TT!E24/'Exchange Rate'!B4-E10"
    # wb.Sheets("A.TT").Cells(14, 5).Value = "=-SUM(TB_GAAP!C:C)"
    wb.Sheets("A.TT").Cells(38, 5).Value = 0


def segment_taxes_based_on_ibit_segmentation(wb, ytd) -> None:
    """
    Gets the segmented IBIT and then allocate taxes based on
    that segmented IBIT.

    :param wb: WorkBook object
    :return: None
    """
    sheet = wb.Worksheets("D1.SEGMENTED_IBIT")
    sheet.Cells(1, 11).Value = "Special\nItems_LC_BT"
    sheet.Cells(1, 12).Value = "State_Tax\nPL sign"
    sheet.Cells(1, 13).Value = "Current_Federal_Tax\nPL sign"
    sheet.Cells(1, 14).Value = "Deferred_Federal_Tax\nPL sign"
    sheet.Cells(1, 15).Value = "Deferred_Tax_GAAP\nPL sign"
    sheet.Cells(1, 16).Value = "Perms_USD"
    sheet.Cells(1, 17).Value = "State_USD"
    sheet.Cells(1, 18).Value = "Federal_USD"

    # this section inputs formulas based on income tax calculated in the first tab
    for i in range(1, sheet.UsedRange.Rows.Count):  # to count how many rows are populated in that sheet
        sheet.Cells(1 + i, 10).Value = """=IF(INDIRECT("RC[1]",FALSE) = 0,\nINDIRECT("RC[-6]",FALSE)\n/(SUM(INDIRECT("C[-6]",FALSE))-SUM(INDIRECT("C[1]",FALSE))/'Exchange Rate'!$B$4),\n0)"""
        sheet.Cells(1 + i, 11).Value = 0 if  sheet.Cells(1 + i, 8).Value != 100 else sheet.Cells(1 + i, 4).Value * wb.Worksheets("Exchange Rate").Cells(4,2).Value
        sheet.Cells(1 + i,
                    12).Value = """=INDIRECT("RC[-1]",FALSE)\n*A.TT!$D$51\n\n+INDIRECT("RC[-2]",FALSE)\n*(A.TT!$E$51*'Exchange Rate'!$B$4\n-(SUM(INDIRECT("C[-1]",FALSE))*A.TT!$D$51))"""
        sheet.Cells(1 + i,
                    13).Value = """=+(INDIRECT("RC[-2]",FALSE)\n*((A.TT!$E$30-A.TT!$D$51)/(1-A.TT!$D$51)))\n\n+(A.TT!$E$42*'Exchange Rate'!$B$4\n-(SUM(INDIRECT("C[-2]",FALSE))*((A.TT!$E$30-A.TT!$D$51)/(1-A.TT!$D$51))))\n*INDIRECT("RC[-3]",FALSE)\n-INDIRECT("RC[-8]",FALSE)"""
        sheet.Cells(1 + i, 14).Value = """=INDIRECT("RC[-4]",FALSE)*A.TT!$E$43*'Exchange Rate'!$B$4"""
        sheet.Cells(1 + i, 15).Value = """=INDIRECT("RC[-5]",FALSE)*A.TT!$E$44*'Exchange Rate'!$B$4-INDIRECT("RC[-9]",FALSE)"""
        sheet.Cells(1 + i, 16).Value = """=-INDIRECT("RC[-6]",FALSE)*A.TT!$E$26/'Exchange Rate'!$B$4"""
        sheet.Cells(1 + i, 17).Value = """=-INDIRECT("RC[-5]",FALSE)/'Exchange Rate'!$B$4"""
        sheet.Cells(1 + i, 18).Value = """=-(INDIRECT("RC[-5]",FALSE)+INDIRECT("RC[-3]",FALSE)+INDIRECT("RC[-4]",FALSE)+INDIRECT("RC[-13]",FALSE)+INDIRECT("RC[-12]",FALSE))/'Exchange Rate'!$B$4"""

        sheet.Cells(1 + i, 5).Value = 0 if ytd == 'ytd' else  sheet.Cells(1 + i, 5).Value
        sheet.Cells(1 + i, 6).Value = 0 if ytd == 'ytd' else  sheet.Cells(1 + i, 6).Value

        sheet.Cells(1 + i, 9).Interior.Color = 65535 if sheet.Cells(1 + i, 9).Value is None else -4142

    # Formating stuff
    sheet.Range("J1").Copy()
    sheet.Range("K1:Q1").PasteSpecial(Paste=-4122,
                                      Operation=-4142,
                                      SkipBlanks=False,
                                      Transpose=False)
    sheet.Columns("K:Q").ColumnWidth = 21.44
    sheet.Range("K1:Q1").WrapText = True
    sheet.Range("K1:Q1").HorizontalAlignment = -4152

    sheet.Columns("J:J").NumberFormat = "0.0%"
    sheet.Columns("K:Q").NumberFormat = "#,##0"

    sheet.Range(f"L2:O{sheet.UsedRange.Rows.Count}").Interior.Color = 14548957
    sheet.Range(f"k2:k{sheet.UsedRange.Rows.Count}").Interior.Color = 16777062

    return None


def generate_pivot_table(file_path, er, country, wb, xl) -> None:
    """
    Generates a pivot table to serve as a mask
    to input data in corptax and EXM.

    :param file_path: file path to be used
    :param er: exchange rate
    :return: None
    """

    ws_data = wb.Worksheets("D1.SEGMENTED_IBIT")
    ws_data.Columns("P:R").EntireColumn.Hidden = True
    ws_report = wb.Worksheets('D.SEGMENTATION_MASK')

    ws_report.Columns("A:K").Delete(Shift=-4159)

    try:

        pt_cache = wb.PivotCaches().Create(1, SourceData=ws_data.Range("B:R"))

        pt = pt_cache.CreatePivotTable(ws_report.Range("A1"), "MyReport")

        # toggle grand totals
        pt.ColumnGrand = False
        pt.RowGrand = True

        # pivot table styles
        pt.TableStyle2 = "PivotStyleLight22"

        field_columns = {}
        field_columns['FS_EXM'] = pt.PivotFields('FS_EXM')

        field_columns['FS_EXM'].Orientation = 2
        field_columns['FS_EXM'].Position = 1

        field_values = {}
        field_values['IBIT_USD HELP(HURT)'] = pt.PivotFields('IBIT_USD HELP(HURT)')
        field_values['State_USD'] = pt.PivotFields('State_USD')
        field_values['Perms_USD'] = pt.PivotFields('Perms_USD')

        field_values['IBIT_USD HELP(HURT)'].Orientation = 4
        field_values['IBIT_USD HELP(HURT)'].Function = -4157
        # field_values['IBIT'].NumberFormat = "#,##0.00"

        field_values['State_USD'].Orientation = 4
        field_values['State_USD'].Function = -4157
        # field_values['State_USD'].NumberFormat = "#,##0.00"

        field_values['Perms_USD'].Orientation = 4
        field_values['Perms_USD'].Function = -4157
        # field_values['Perms_USD'].NumberFormat = "#,##0.00"

        ws_report.PivotTables("MyReport").PivotFields("FS_EXM").PivotItems("(blank)").Visible = False

        ws_report.PivotTables("MyReport").DataPivotField.Orientation = 1
        ws_report.PivotTables("MyReport").DataPivotField.Position = 1

        ws_report.PivotTables("MyReport").CalculatedFields().Add(Name="Taxable_Income",
                                                                 Formula="=Perms_USD+'IBIT_USD HELP(HURT)' +State_USD",
                                                                 UseStandardFormula=True)

        pt.AddDataField(pt.PivotFields('Taxable_Income'))

        field_values['Federal_USD'] = pt.PivotFields('Federal_USD')
        field_values['Federal_USD'].Orientation = 4
        field_values['Federal_USD'].Function = -4157
        # field_values['Federal_USD'].NumberFormat = "#,##0.00"

        ws_report.PivotTables("MyReport").CalculatedFields().Add(Name="Earnings_AT",
                                                                 Formula="=Federal_USD+State_USD+'IBIT_USD HELP(HURT)'",
                                                                 UseStandardFormula=True)

        pt.AddDataField(pt.PivotFields('Earnings_AT'))

        # pt.PivotFields('Taxable_Income').Orientation = 1
        # pt.PivotFields('Taxable_Income').Function = -4157
        # pt.PivotFields('Taxable_Income').NumberFormat = "#,##0.00"

        # ws_report.PivotTables("MyReport").PivotFields("Taxable_Income").Orientation = 1

        ws_report.Cells(1, 2).Value = "Segment"
        ws_report.Cells(3, 1).Value = "IBIT "
        ws_report.Cells(4, 1).Value = "State Tax"
        ws_report.Cells(5, 1).Value = "Perms "
        ws_report.Cells(6, 1).Value = "Taxable Income "
        ws_report.Cells(7, 1).Value = "Federal Tax "
        ws_report.Cells(8, 1).Value = "Earnings After Tax "

        ws_report.Range("B:AB").NumberFormat = "#,##0.0,,"

        ws_report.Range("B2:AB2").NumberFormat = "@"
        ws_report.Range("B2:AB2").WrapText = True
        ws_report.Range("B2:AB2").HorizontalAlignment = -4108
        ws_report.PivotTables("MyReport").GrandTotalName = "Corporate Total"

    except Exception as e:
        print("Warning error due to: ", e)
        pass


def paste_plot(xl, wb):
    sheet = wb.Sheets('A.TT')
    wb.Worksheets('A.TT').Activate()
    sheet.Cells(54, 2).Value = "'+ TOTAL TAX USD"
    top = sheet.Cells(58, 2).Top;
    left = sheet.Cells(58, 2).Left  # To get the Top & Left coordinates

    xl.ActiveSheet.Shapes.AddPicture(f'{os.environ["USERPROFILE"]}\\figure.png', False, True, left,
                                     top,
                                     -1, -1)

    sheet.Cells(1, 1).Select()
