"""
File related to the program GUI
using mainly tkinter facilities to
display the user interface
"""
# Import section
import os
import fed
import excel_manipulator
import spartanASCII
import tkinter as tk
import pandas as pd
import datetime as dt
import gui_styles as styles
from snowflake_aid import SnowSQL
from tkinter import filedialog



#icon stuff
x = dt.date.today()

y = x.isoformat()[:-2] + '01'
z = dt.date.fromisoformat(y)
year = z - dt.timedelta(days=1)
month = year.month
year = year.year


class MainGUI(tk.Tk):
    """
    Class to encapsulate all properties and functions to the main window of the program.
    """

    def __init__(self,email):
        """
        Initialize the program.
        """
        super().__init__()
        #os.system('cls')
        print(spartanASCII.BANDEIRA_COXA)

        fed.variables.create_template_adjustment_table()
        self.create_folder_on_desktop()
        print("Logging credentials on Snowflake")
        self.email = email
        self.__FED = SnowSQL(self.email)
        self.adjustment_path = None
        self.wm_iconbitmap(fed.variables.generate_icon(styles.icon))
        self.geometry("1150x350")
        self.title('SPARTA Tool v1.0.5')
        self.upperFrame = tk.Frame(self, bg='green')
        self.upperFrame.pack(fill="both", side=tk.TOP)
        self.Label1 = tk.Label(self.upperFrame, text="Streamlined Tax Tool", bg='green', fg='white',
                               font=styles.header)
        self.Label1.pack(side='top')

        """
        ******************* F R A M E S ***********************
        """
        self.second_frame1 = tk.Frame(self)
        self.second_frame1.pack(fill='both', side=tk.TOP)

        self.random_frame1 = tk.Frame(self)
        self.random_frame1.pack(fill='both', side=tk.TOP)

        """
        Text variables Sections     
        """
        self.ru = tk.StringVar()
        self.ru.set('1556')
        self.year = tk.StringVar()
        self.year.set(str(year))
        self.period = tk.StringVar()
        self.period.set(str(month) if month > 9 else f'0{str(month)}')
        self.signal = False


        bgcolor = 'blue'
        fgcolor = 'black'

        l2 = tk.Label(self, text="Insert RU:", font=styles.header, fg=fgcolor)
        l2.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=9, font=styles.body, textvariable=self.ru, fg=bgcolor).pack(in_=self.second_frame1,
                                                                                         side=tk.LEFT)

        l3 = tk.Label(self, text="Year:", font=styles.header, fg=fgcolor)
        l3.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=10, font=styles.body, textvariable=self.year, fg=bgcolor).pack(in_=self.second_frame1,
                                                                                            side=tk.LEFT)

        l4 = tk.Label(self, text="Period:", font=styles.header, fg=fgcolor)
        l4.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=10, font=styles.body, textvariable=self.period, fg=bgcolor).pack(in_=self.second_frame1,
                                                                                              side=tk.LEFT)

        l8 = tk.Label(self, text="YTD or SA:", font=styles.header, fg=fgcolor)
        l8.pack(in_=self.second_frame1, side=tk.LEFT)
        cm2 = ["YTD", "SA"]
        self.ytd = tk.StringVar()
        self.ytd.set("ytd")
        drop = tk.OptionMenu(self, self.ytd, *cm2)
        drop.pack(in_=self.second_frame1, side=tk.LEFT)

        l7 = tk.Label(self, text="Method:", font=styles.header, fg=fgcolor)
        l7.pack(in_=self.second_frame1, side=tk.LEFT)
        cm = ["Local", "GAAP"]
        self.methodology = tk.StringVar()
        self.methodology.set("Local")
        drop = tk.OptionMenu(self, self.methodology, *cm)
        drop.pack(in_=self.second_frame1, side=tk.LEFT)

        tk.Label(text="=========================================================================================",
                 font=styles.body).pack(in_=self.random_frame1, side=tk.LEFT)

        tk.Button(text="RUN !", font=styles.header, command=self.run).place(x=200, y=100)
        tk.Button(text="MASSIVE", font=styles.header, command=self.massive_run).place(x=400, y=100)

        tk.Label(text="=========================================================================================",
                 font=styles.body).place(x=1, y=250)

        tk.Label(
            text="SPARTA: Streamlined Process to Account & Report Tax Amounts\n  This tool was designed to "
                 "automatically calculate tax massively with FED data.",
            font=styles.body).place(x=140, y=270)

        tk.Label(
            text=spartanASCII.spartanASCII,
            font=styles.ascii).place(x=800, y=35)

    def run(self):
        """Runs the program for an single Reporting Unit using method
        "run_individual_ru" below."""
        NOW = dt.datetime.now()
        self.adjustment_path = fed.variables.get_path()
        ru = self.ru.get()
        r_year = self.year.get()
        period = self.period.get()
        methodology = self.methodology.get()

        self.run_individual_ru(ru, r_year, period, methodology, self.ytd.get())

        print(f'Total execution time : {dt.datetime.now()-NOW}')


    def run_individual_ru(self, ru: str, r_year: str, period: str, methodology: str, ytd :str) -> None:
        """
        Run the program for a single reporting unit.

        :rtype: None -> generates an excel file
        :param ru: reporting unit to be run
        :param r_year: the reporting year
        :param period: the period, meaning month no.
        :param methodology: type int methodology used, if using local books or ExxonBooks
        :param ytd: string telling if it is a YTD Calc or not
        """

        try:
            country, tax_rate, currency = self.get_master(ru=ru)
        except Exception as e:
            print(e)
            return

        self.signal = False
        file_name = f'{ru}-{period}-{r_year}--Generated-by-Sparta.xlsx'
        file_path = f"{os.environ['USERPROFILE']}\\Desktop\\Sparta_Output\\{file_name}"
        trial_balance_GAAP, trial_balance_local, trial_balance_delta_between_GAAPs, perms_from_excel, exchange_rate, \
            credits_from_excel, GTD_detail, temps_from_excel, signal = self.get_data_from_fed(
                ru, r_year, period, currency, tax_rate, methodology, ytd)

        segmentation, tidy_segmentation = fed.get_segmentation_from_fed(self.__FED, ru, r_year, period, ytd)

        if len(trial_balance_GAAP) != 0:
            self.signal = True

        df_dictionary = {"exchange_rate": exchange_rate, "trial_balance_GAAP": trial_balance_GAAP,
                         "trial_balance_local": trial_balance_local,
                         "trial_balance_delta_between_GAAPs": trial_balance_delta_between_GAAPs,
                         "perms_from_excel": perms_from_excel, "credits_from_excel": credits_from_excel,
                         "temps_from_excel": temps_from_excel,
                         "GTD_detail": GTD_detail,
                         "segmentation": segmentation,
                         "tidy_segmentation": tidy_segmentation}
        excel_manipulator.pandas_excel(file_name=file_name, dataframe_dictionary=df_dictionary)

        xl, wb = excel_manipulator.open_excel_file(file_path)
        xl.DisplayAlerts = False

        # Sheet creation

        excel_manipulator.create_total_tax_mask(xl, wb, dataframe_dictionary=df_dictionary, tax_rate=tax_rate,
                                                methodology=methodology, country=country)
        excel_manipulator.create_cover(wb, ru, r_year, period, methodology)
        wb.Worksheets(1).Name = 'Cover'
        wb.Worksheets(2).Name = 'A.TT'

        excel_manipulator.create_cf216_mask(xl=xl, wb=wb, dataframe_dictionary=df_dictionary,
                                            currency=currency)

        # Sheet Names

        wb.Worksheets(4).Name = 'C.CF216'

        excel_manipulator.create_memo_forex(wb)

        wb.Sheets("C1.MEMO FOREX").Move(Before=wb.Worksheets(5))




        excel_manipulator.segment_taxes_based_on_ibit_segmentation(wb=wb, ytd= ytd)

        wb.SaveAs(f"{file_path}", ConflictResolution=2)
        wb.Close()
        er = exchange_rate['Close']

        #excel_manipulator.generate_pivot_mask(file_path =  file_path, er = er, country = country)


        # reopens the file because of the generate pivot mask uses writer instead of PYWIN32
        xl, wb = excel_manipulator.open_excel_file(file_path)
        xl.DisplayAlerts = False


        excel_manipulator.generate_pivot_table(file_path, er, country, wb,xl)



        excel_manipulator.format_cells(xl, wb)
        excel_manipulator.format_tab_color(wb)
        if country == 'US':
            excel_manipulator.transform_in_us(wb)

        if signal:
            excel_manipulator.paste_plot(xl, wb)



        wb.SaveAs(f"{file_path}", ConflictResolution=2)
        wb.Close()

        print(f"File successfully saved at : {file_path}")

    def massive_run(self):
        """Runs the program for a queue list of Reporting Units defined in
        a separate excel file, by looping that list through method "run"
        above."""
        self.adjustment_path = fed.variables.get_path()
        now = dt.datetime.now()

        filepath = self.get_filenames()
        df = pd.read_excel(filepath[0], dtype=str)
        error_list = []
        for i in df.index:
            temp = df.iloc[i]
            RU = temp[0]
            r_year = temp[1]
            period = temp[2]
            methodology = temp[3]
            ytd = temp[4]

            r_list = [RU, r_year, period, methodology]

            try:
                self.run_individual_ru(RU, r_year, period, methodology,ytd)
                print(self.signal)
            except Exception as e:
                r_list.append(str(e))
                error_list.append(r_list)

        print("MASSIVE JOB FINISHED")
        if len(error_list) != 1:
            print("Errors processing RUs to follow:")
            for i in error_list:
                print(i)
        later = dt.datetime.now()
        print(dt.datetime.strftime(now, '%B %dth, %Y  %H:%M:%S'), "+++",
              dt.datetime.strftime(later, '%B %dth, %Y  %H:%M:%S'))

        print(f"Elapsed time {later - now}")

    def get_data_from_fed(self, ru: str, r_year: str, period: str, currency: str, tax_rate: float, method: str,ytd: str):
        """To retrieve data from FED and return Trial Balances plus Perms
        from MS Excel stored in the lan folder.

        :param ru: reporting unit to be run
        :param r_year: the reporting year
        :param period: the period, meaning month no.
        :param currency: local currency of ru, used to revaluate
         tax at USD
        :param tax_rate: country tax rate
        :param ytd: string telling if it is a YTD Calc or not

        :rtype:tuple[
            pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame,
            pd.DataFrame, pd.DataFrame, pd.DataFrame]"""

        print("GAAP Trial Balance:")
        trial_balance_GAAP, signal = fed.get_trial_balance_gaap(self.__FED, ru, r_year, period,tax_rate)
        print("Local Trial Balance:")
        trial_balance_local = fed.get_trial_balance_local(self.__FED, ru, r_year, period)
        print("GAAP & Local Trial Balance:")
        trial_balance_delta_between_GAAPs, exchange_rate, GTD_detail = \
            fed.get_trial_balance_delta_between_gaaps_PL(
                self.__FED, ru, r_year, period, currency, tax_rate)\
            if method == 'Local' else fed.get_trial_balance_delta_between_gaaps_BS(
                self.__FED, ru, r_year, period, currency, tax_rate)
        print("Reading Perms from excel...")
        perms_from_excel = fed.get_perms_from_excel(ru, r_year, period, self.adjustment_path)

        print("Reading Temps from excel...")
        temps_from_excel = fed.get_temps_from_excel(ru, r_year, period, self.adjustment_path)

        print("Reading Credits from excel...")
        credits_from_excel = fed.get_credits_from_excel(ru, r_year, period, self.adjustment_path)

        return trial_balance_GAAP, trial_balance_local, trial_balance_delta_between_GAAPs, perms_from_excel, \
            exchange_rate, credits_from_excel, GTD_detail, temps_from_excel, signal




    @staticmethod
    def create_folder_on_desktop():
        """Creates folder on desktop named Sparta_Output, if not exists yet.
        This folder will group all files created as a result of this program."""
        p = f'{os.environ["USERPROFILE"]}\\Desktop\\Sparta_Output'
        if not os.path.isdir(p):
            os.system(f'mkdir {p}')

    @staticmethod
    def get_filenames() -> list[str]:
        """Function to call a filedialog to select excel file
        that contains the "massive_run" function inputs."""
        filenames = filedialog.askopenfilenames(
            initialdir=os.environ["USERPROFILE"], title="Select a File",
            filetypes=(("Excel", "*.xl*"), ("all files", "*.*")))

        return filenames

    def get_master(self, ru):
        self.adjustment_path = fed.variables.get_path()
        master = pd.read_excel(io=self.adjustment_path, sheet_name='RU_Master', dtype={'RU': str})
        try:
            country, tax_rate, currency = master[master['RU'] == ru].iloc[0].to_list()[1:]
        except IndexError:
            print(f'ERROR :RU not found in file {self.adjustment_path}', f'error{IndexError}')
            return
        except Exception as e:
            print(f'Error due to {e}')
            return

        print(f'Reporting Unit country :{country}, currency {currency}, tax_rate {tax_rate}')

        return country, tax_rate, currency






