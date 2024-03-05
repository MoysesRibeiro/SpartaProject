import os
import numpy as np
import pandas as pd
import variables
import warnings
import datetime as dt
import yfinance as yf
import segment_table
from variables import __SCALLING_FACTOR__
from io import StringIO
import matplotlib.pyplot as plt

warnings.filterwarnings('ignore')


def get_trial_balance_gaap_for_chart(snow_flake_connection, ru, year, period, STR):
    query = f'''
            SELECT "G/L Acct", sum("GC Amount")/{__SCALLING_FACTOR__} AS AMOUNT, 
                sum("LC Amount")/{__SCALLING_FACTOR__} AS LC_AMOUNT,
                "Period"

            FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_FI_TRIAL_BALANCE_REPORT
            WHERE "Company Code" ='{ru}'
                AND "Fiscal Yr" ='{year}'
                AND "Period" <= '{period}'
                AND LENGTH("G/L Acct") = '8'
                AND RIGHT(LEFT("G/L Acct",4),2) <> '98'
            GROUP BY "G/L Acct","Period"
            ORDER BY "G/L Acct", "Period"
            '''

    df = snow_flake_connection.execute_query(query, ru, year, period).fetch_pandas_all()
    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df['AMOUNT'] = np.round(df['AMOUNT'], 0)
    df['LC_AMOUNT'] = df['LC_AMOUNT'].astype(np.int64)
    df['LC_AMOUNT'] = np.round(df['LC_AMOUNT'], 0)

    df['Type'] = None
    df['Period'] =  df['Period'].astype(int)
    for i in range(len(df['G/L Acct'])):
        df['Type'][i] = 'IBIT' if str(df['G/L Acct'][i])[:2] not in ("89", "90") else "Tax"

    tax = df[df['Type'] == "Tax"]
    ibit = df[df['Type'] == "IBIT"]

    ibit = ibit[["Period", 'AMOUNT']].groupby("Period").sum() * -1
    tax = tax[["Period", 'AMOUNT']].groupby("Period").sum()

    ibit['TAX'] = tax

    df = ibit.assign(IBIT_YTD=ibit['AMOUNT'].cumsum())
    df = df.assign(TAX_YTD=ibit['TAX'].cumsum())

    df['ETR'] = df['TAX_YTD'] / df['IBIT_YTD'] * 100
    df['STR'] = STR
    df = df.iloc[:-1]

    df[df.columns[:4]] = np.round(df[df.columns[:4]] / 1000000, 0)
    plt.style.use('grayscale')

    fig = plt.figure(figsize=(10, 6), dpi=100)

    fig.set_facecolor('white')
    ax1 = fig.add_axes(rect=(0, 0.7, 1, 0.7))
    plt.xticks(df.index)
    ax2 = fig.add_axes(rect=(0, 0.3, 1, .3))
    plt.xticks(df.index)

    ax1.set_title("IBIT & Tax Analysis", size=14)
    ax1.grid(True, alpha=0.05)
    ax1.plot(df['IBIT_YTD'], alpha=0.3, label='IBIT')

    scale = df['IBIT_YTD'].iloc[-1] / 50
    for i, txt in enumerate(np.round(np.round(df['IBIT_YTD'].to_list(), 1), 1), 1):
        ax1.annotate("${:,}".format(int(txt)), (df.index[i - 1], df['IBIT_YTD'][i] + scale))

    ax1.plot(df['TAX_YTD'], label='Tax', color='green')

    for i, txt in enumerate(np.round(np.round(df['TAX_YTD'].to_list(), 1), 1), 1):
        nbr = 0 if np.isnan(txt) else txt
        ax1.annotate("${:,}".format(int(nbr) if nbr is not np.nan else 0), (df.index[i - 1], df['TAX_YTD'][i] + scale))

    ax1.legend(loc='best', fontsize=14)

    lista = []
    for i in ax1.get_yticks():
        lista.append("${:,}".format(int(i)))

    ax1.set_yticklabels(lista)

    ax2.set_title("Effective Tax Rate", size=14)
    ax2.plot(df['STR'], alpha=0.3)
    ax2.scatter(df['ETR'].index, df['ETR'], label="ETR", marker="+", s=115, color='green')
    ax2.set_xlabel("Period", size=14)
    ax2.grid(True, alpha=0.05)

    lista = []
    for i in ax2.get_yticks():
        lista.append("%{:,}".format(int(i)))

    ax2.set_yticklabels(lista)

    for i, txt in enumerate(np.round(df['ETR'].to_list(), 1), 1):
        ax2.annotate(str(txt), (df.index[i - 1] + 0.2, df['ETR'][i] - 0.1))

    plt.savefig(f"{os.environ['USERPROFILE']}\\figure.png",bbox_inches = 'tight')

    return True

def get_trial_balance_gaap(snow_flake_connection, ru, year, period, STR):
    query = f'''
        SELECT "G/L Acct", sum("GC Amount")/{__SCALLING_FACTOR__} AS AMOUNT, 
            sum("LC Amount")/{__SCALLING_FACTOR__} AS LC_AMOUNT
            
        FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_FI_TRIAL_BALANCE_REPORT
        WHERE "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LENGTH("G/L Acct") = '8'
            AND RIGHT(LEFT("G/L Acct",4),2) <> '98'
        GROUP BY "G/L Acct"
        ORDER BY "G/L Acct"
        '''

    df = snow_flake_connection.execute_query(query, ru, year, period).fetch_pandas_all()
    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df['AMOUNT'] = np.round(df['AMOUNT'], 0)
    df['LC_AMOUNT'] = df['LC_AMOUNT'].astype(np.int64)
    df['LC_AMOUNT'] = np.round(df['LC_AMOUNT'], 0)

    signal = False

    try:
        signal = get_trial_balance_gaap_for_chart(snow_flake_connection, ru, year, period, STR)
    except Exception as e:
        print(e)
    df.to_excel(f"{os.environ['USERPROFILE']}\\SpartaTemp_GAAP.xlsx")

    return df,signal


def get_trial_balance_local(snow_flake_connection, ru, year, period):
    query = f'''
        SELECT "G/L Acct", sum("LC Amount")/{__SCALLING_FACTOR__} AS AMOUNT, 
            sum("GC Amount")/{__SCALLING_FACTOR__} AS GC_AMOUNT
        FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_FI_TRIAL_BALANCE_REPORT
        WHERE "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LENGTH("G/L Acct") = '8'
            AND RIGHT("G/L Acct",3) <> '799'
        OR "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LEFT("G/L Acct",2) = 'N0'
        GROUP BY "G/L Acct"
        ORDER BY "G/L Acct"
        '''

    df = snow_flake_connection.execute_query(query, ru, year, period).fetch_pandas_all()
    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df['AMOUNT'] = np.round(df['AMOUNT'], 0)
    df['GC_AMOUNT'] = df['GC_AMOUNT'].astype(np.int64)
    df['GC_AMOUNT'] = np.round(df['GC_AMOUNT'], 0)
    df['MAIN'] = df['G/L Acct'].str[:2]
    df.to_excel(f"{os.environ['USERPROFILE']}\\SpartaTemp_local.xlsx")
    return df

def scan_excahnge_rate(currency: str, data: type(dt.datetime)) -> object:
    """
    Gets the exchange rate from yahoo finance if available.

    :param currency: which currency / local currency
    :param data: date in which the report needs to be updated
    :returns: an exchange rate series and er (individual exchange rate)
    """

    try:
        print("Scanning for the exchange rate...")
        exchange_rate = \
            yf.Ticker(f'{currency}=X').history(start=str((data - dt.timedelta(days=30)).date()),
                                               end=str(data.date())).iloc[
                -1]
        er = exchange_rate['Close']
        exchange_rate.name = (str(exchange_rate.name)[:10])
        exchange_rate = exchange_rate[1:4]

    except Exception as e:
        fx = [1, 1, 1, 1]
        exchange_rate = pd.Series(fx, name='USD:USD', index=['Open', 'High', 'Low', 'Close'])

        er = 1
        print("Couldn't download exchange rate, due to ", e)

    return exchange_rate, er

def get_trial_balance_delta_between_gaaps_PL(snow_flake_connection, ru, year, period, currency, tax_rate):


    data = variables.get_end_of_month_date(int(year), int(period))

    exchange_rate, er = scan_excahnge_rate(currency, data)

    query = f'''
        SELECT "G/L Acct", sum("GC Amount")/{__SCALLING_FACTOR__} AS AMOUNT, 
            sum("LC Amount")/{__SCALLING_FACTOR__} AS LC_AMOUNT
        FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_FI_TRIAL_BALANCE_REPORT
        WHERE "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LENGTH("G/L Acct") = '8'
            AND RIGHT("G/L Acct",3) = '799'
        OR "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LENGTH("G/L Acct") = '8'
            AND RIGHT(LEFT("G/L Acct",4),2) = '98'
        OR "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LEFT("G/L Acct",2) = 'N0'
        GROUP BY "G/L Acct"
        ORDER BY "G/L Acct"
        '''

    df = snow_flake_connection.execute_query(query, ru, year, period).fetch_pandas_all()

    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df['LC_AMOUNT'] = df['LC_AMOUNT'].astype(np.int64)
    df["LC_AMT_USD"] = df['LC_AMOUNT'] / er
    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df = np.round(df, 0)

    # manipulation
    df['Is_EMOnly'] = (df["G/L Acct"].str[-3:]) == '799'
    df['Is_Local'] = df['Is_EMOnly'] != True

    GTD_detail = df
    listagem = [str(x) for x in range(89)]
    listagem.append("N0")

    # ver depois
    df['Main'] = ''

    for i in range(0,len(df['Main'])):
        if df["G/L Acct"][i][:2] != 'N0':
            df['Main'][i] = df["G/L Acct"][i][:2]

        else:
            df['Main'][i] = df["G/L Acct"][i][2:4]


    GTD_detail = GTD_detail[GTD_detail['Main'].isin(listagem)]
    GTD_detail = GTD_detail[GTD_detail['G/L Acct'].str[:4] != 'N089']
    GTD_detail = GTD_detail[GTD_detail['G/L Acct'].str[:4] != 'N090']
    df = df[df['G/L Acct'] != '21959799']
    df = df[df['G/L Acct'] != '71959799']


    pivot_local = df[df['Is_Local'] == True].groupby('Main').sum()['LC_AMT_USD']
    pivot_local.name = 'Local'
    pivot_GAAP = df[df['Is_EMOnly'] == True].groupby('Main').sum()['AMOUNT']
    pivot_GAAP.name = 'GAAP'

    merged = pd.DataFrame(data=[pivot_GAAP, pivot_local]).T
    merged = merged.fillna(0)

    merged['GTD (Delta)'] = merged['GAAP'] - merged['Local']

    merged = merged[~merged.index.isin(['89', '90', '95', '96'])]

    merged = merged.sort_values('Main')
    merged['TAX'] = merged['GTD (Delta)'] * -tax_rate / 100
    merged['TAX_LC'] = merged['TAX'] * er

    merged['CONCEPT'] = ''
    for i in range(len(merged['CONCEPT'])):
        merged['CONCEPT'][i] = variables.dictionary_main_vs_concept.get(int(merged.index[i]))



    merged.set_index([merged.index, merged['CONCEPT']], inplace=True)
    merged.drop('CONCEPT', axis=1, inplace=True)
    merged.to_excel(f"{os.environ['USERPROFILE']}\\SpartaTemp_delta_between_GAAPs.xlsx")
    return merged, exchange_rate, GTD_detail

def get_deferred_tax_balances(snow_flake_connection, ru, year, period):
    """
    Gets 465 and 710 accounts balances.

    :param snow_flake_connection: SnowSQL connection instance
    :param ru: Reporting unit
    :param year: year
    :param period: period
    :return: dataframe with the balances
    """

    query_dt = f'''
            SELECT "G/L Acct", ROUND(sum("GC Amount"),2)AS AMOUNT, ROUND(sum("LC Amount"),2) AS LC_AMOUNT
            FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_FI_TRIAL_BALANCE_REPORT
            WHERE "Company Code" ='{ru}'
                AND "Fiscal Yr" ='{int(year)-1}'
                AND "Period" <= '{12}'
                AND LENGTH("G/L Acct") = '9'
                AND "G/L Acct" LIKE ANY ('465%','710%')

            GROUP BY "G/L Acct"
            ORDER BY "G/L Acct"
            '''
    df = snow_flake_connection.execute_query(query_dt, ru, year, period).fetch_pandas_all()

    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df['LC_AMOUNT'] = df['LC_AMOUNT'].astype(np.int64)
    # df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df = np.round(df, 0)

    df['Base'] = ''

    for i in range(0,len(df['Base'])):
        if df["G/L Acct"][i][:1] != 'N':
            df['Base'][i] = df["G/L Acct"][i][:6]

        else:
            df['Base'][i] = df["G/L Acct"][i][1:7]

    df['CONCEPT'] = ''
    for i in range(len(df['CONCEPT'])):
        df['CONCEPT'][i] = variables.dictionary_base_vs_concept.get(int(df['Base'][i]))

    return df

def get_trial_balance_delta_between_gaaps_BS(snow_flake_connection, ru, year, period, currency, tax_rate):
    data = variables.get_end_of_month_date(int(year), int(period))

    exchange_rate, er = scan_excahnge_rate(currency, data)

    query = f'''
        SELECT "G/L Acct", sum("GC Amount")/{__SCALLING_FACTOR__} AS AMOUNT, 
            sum("LC Amount")/{__SCALLING_FACTOR__} AS LC_AMOUNT
        FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_FI_TRIAL_BALANCE_REPORT
        WHERE "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LENGTH("G/L Acct") = '9'
            AND RIGHT("G/L Acct",3) = '799'
        OR "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LENGTH("G/L Acct") = '9'
            AND RIGHT(LEFT("G/L Acct",5),2) = '98'
        OR "Company Code" ='{ru}'
            AND "Fiscal Yr" ='{year}'
            AND "Period" <= '{period}'
            AND LEFT("G/L Acct",1) = 'N'
            AND LEFT("G/L Acct",2) <> 'N0'
        GROUP BY "G/L Acct"
        ORDER BY "G/L Acct"
        '''



    df = snow_flake_connection.execute_query(query, ru, year, period).fetch_pandas_all()
    df_dt = get_deferred_tax_balances(snow_flake_connection, ru, year, period)

    df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df['LC_AMOUNT'] = df['LC_AMOUNT'].astype(np.int64)
    #df['AMOUNT'] = df['AMOUNT'].astype(np.int64)
    df = np.round(df, 0)

    # manipulation
    df['Is_EMOnly'] = (df["G/L Acct"].str[-3:]) == '799'
    df['Is_Local'] = df['Is_EMOnly'] != True

    GTD_detail = df

    # ver depois
    df['Main'] = ''

    for i in range(0,len(df['Main'])):
        if df["G/L Acct"][i][:1] != 'N':
            df['Main'][i] = df["G/L Acct"][i][:3]

        else:
            df['Main'][i] = df["G/L Acct"][i][1:4]



    pivot_local = df[df['Is_Local'] == True].groupby('Main').sum()['LC_AMOUNT']
    pivot_local.name = 'Local'
    pivot_GAAP = df[df['Is_EMOnly'] == True].groupby('Main').sum()['LC_AMOUNT']
    pivot_GAAP.name = 'GAAP'

    merged = pd.DataFrame(data=[pivot_GAAP, pivot_local]).T
    merged = merged.fillna(0)

    merged['GTD (Delta)'] = merged['GAAP'] - merged['Local']

    merged = merged[~merged.index.isin(['609', '710', '810', '841','845','846','860','901','999'])]

    merged = merged.sort_values('Main')
    merged['TAX'] = merged['GTD (Delta)'] * -tax_rate / 100
    #merged['TAX_USD'] = merged['TAX'] / er

    merged['CONCEPT'] = ''
    for i in range(len(merged['CONCEPT'])):
        merged['CONCEPT'][i] = variables.dictionary_main_vs_concept.get(int(merged.index[i]))


    merged_detail = merged
    merged.reset_index(inplace=True)


    merged = pd.concat(objs = [merged,df_dt[['CONCEPT','LC_AMOUNT']]], axis = 0, join = 'outer')

    merged = merged.groupby('CONCEPT').sum()[["GAAP", "Local", "GTD (Delta)", "TAX","LC_AMOUNT"]]
    merged.rename(columns={"LC_AMOUNT":"PRIOR_YEAR"},inplace=True)
    merged['POSTING'] = merged['TAX'] - merged['PRIOR_YEAR']

#
#    merged = merged.concat(other = df_dt, on = 'CONCEPT', how = 'left')

    df_dt.to_excel(f"{os.environ['USERPROFILE']}\\{ru}-Detail_of_DEFERRED_BALANCESv2.xlsx")
    merged_detail.to_excel(f"{os.environ['USERPROFILE']}\\{ru}-Detail_of_DEFERRED_BALANCES.xlsx")


    return merged, exchange_rate, GTD_detail

def get_perms_from_excel(ru, year, period, p):
    print("Getting perms from :", p)
    df = pd.read_excel(p, sheet_name="Perms", dtype=str)
    df['PERM_AMOUNT'] = df['PERM_AMOUNT'].astype(np.float64)
    df = df[df['YEAR'] == year]
    df = df[df['PERIOD'] == period]
    df = df[df['RU'] == ru]
    df.to_excel(f"{os.environ['USERPROFILE']}\\SpartaTemp_Perms.xlsx")
    return df

def get_credits_from_excel(ru, year, period, p):
    df = pd.read_excel(p, sheet_name="Credits", dtype=str)
    df['CR_AMOUNT'] = df['CR_AMOUNT'].astype(np.float64)
    df = df[df['YEAR'] == year]
    df = df[df['PERIOD'] == period]
    df = df[df['RU'] == ru]
    df.to_excel(f"{os.environ['USERPROFILE']}\\SpartaTemp_Credits.xlsx")
    return df

def get_temps_from_excel(ru, year, period, p):
    df = pd.read_excel(p, sheet_name="Temps", dtype=str)
    df['TEMP_AMOUNT'] = df['TEMP_AMOUNT'].astype(np.float64)
    df = df[df['YEAR'] == year]
    df = df[df['PERIOD'] == period]
    df = df[df['RU'] == ru]
    df.to_excel(f"{os.environ['USERPROFILE']}\\SpartaTemp_Temps.xlsx")
    return df

def get_segmentation_from_fed(snow_flake_connection, ru, year, period, ytd):
    """
    Gets Ibit segmented by BA and WWPC from table
    FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL.

    :param snow_flake_connection: Snowflake connection
    :param ru: reporting unit
    :param year: year
    :param period: Period string of two caracters, example 09 or 12
    :param ytd: string telling if it is a YTD Calc or not
    :return: None
    """
    print('Querying segmentation from Special Ledger: ')

    if ytd == 'ytd':
        query = f'''
                 SELECT "Bus Area",WWPC , sum("GC Amount") AS IBIT
                     FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL
                 WHERE "Company Code" = '{ru}'
                     AND "Fiscal Yr" = '{year}'
                     AND "Period" <= '{period}'
                     AND LEFT("G/L Acct",2) = '00'
                     AND LEFT("G/L Acct",2) <> 'N0'
                     AND RIGHT(LEFT("G/L Acct",6),2) <> '98'
                     AND LEFT("G/L Acct",4) <> '0089'
                     AND LEFT("G/L Acct",4) <> '0090'
                     AND "Record Type" IN ('0','2') 
                    AND "Source Table" ='YWS01T'
                 
                 GROUP BY "Bus Area", WWPC
                 ORDER BY "Bus Area", WWPC
             '''
    else:
        query = f'''
                         SELECT "Bus Area",WWPC ,
                                SUM(CASE WHEN LEFT("G/L Acct",4) NOT IN ('0089','0090') THEN "GC Amount" ELSE 0 END) AS IBIT,
                                SUM(CASE WHEN LEFT("G/L Acct",4) = '0089' THEN "LC Amount" ELSE 0 END) AS CIT_LC_PLSIGN,
                                SUM(CASE WHEN LEFT("G/L Acct",4) = '0090' THEN "LC Amount" ELSE 0 END) AS DIT_LC_PLSIGN
                         FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL
                         WHERE "Company Code" = '{ru}'
                             AND "Fiscal Yr" = '{year}'
                             AND "Period" <= '{period}'
                             AND LEFT("G/L Acct",2) = '00'
                             AND LEFT("G/L Acct",2) <> 'N0'
                             AND RIGHT(LEFT("G/L Acct",6),2) <> '98'

                             AND "Record Type" IN ('0','2') 
                            AND "Source Table" ='YWS01T'

                         GROUP BY "Bus Area", WWPC
                         ORDER BY "Bus Area", WWPC
                     '''

    df = snow_flake_connection.execute_query(query, ru, year, period).fetch_pandas_all()
    df['IBIT'] = df['IBIT'].astype(np.int64)
    df['IBIT'] = -np.round(df['IBIT'], 0)
    if ytd != 'ytd':
        df['CIT_LC_PLSIGN'] = df['CIT_LC_PLSIGN'].astype(np.int64)
        df['CIT_LC_PLSIGN'] = -np.round(df['CIT_LC_PLSIGN'], 0)
        df['DIT_LC_PLSIGN'] = df['DIT_LC_PLSIGN'].astype(np.int64)
        df['DIT_LC_PLSIGN'] = -np.round(df['DIT_LC_PLSIGN'], 0)
    #df.to_excel(f"{os.environ['USERPROFILE']}\\SpartaSEG_GAAP.xlsx")
    tidy_df = df[df['WWPC'] != '']  # to filter out blank_values


    segment_table_df = pd.read_csv(StringIO(segment_table.segment_table_csv), sep=',')
    pcmapping_df = pd.read_csv(StringIO(segment_table.pc_mapping_wwpc), sep=',')

    tidy_df = pd.merge(tidy_df, segment_table_df, on='WWPC', how='left')

    pcmapping_df = pcmapping_df[pcmapping_df['RU'].isin([ru, 'ALL'])]
    tidy_df = pd.merge(tidy_df, pcmapping_df[["WWPC", "Profit Center"]], on='WWPC', how='left')
    tidy_df.drop_duplicates(['Bus Area','WWPC','IBIT'], inplace=True)

    tidy_df = tidy_df[tidy_df['IBIT'] != 0]
    tidy_df['FS_EXM'].fillna('ALLOCATE_MANUALLY', inplace=True)
    #tidy_df.dropna(subset=['FS_EXM'], inplace=True)

    #tidy_df['Allocation'] = tidy_df['IBIT'] / tidy_df['IBIT'].sum()
    tidy_df['Allocation'] = 0
    tidy_df.rename(columns={'IBIT':'IBIT_USD HELP(HURT)'},inplace=True)
    if ytd != 'ytd':
        tidy_df.drop(labels=['EffDate','BA'], axis='columns', inplace=True)

    return df, tidy_df