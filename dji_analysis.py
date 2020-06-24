# dji_analysis.py

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf
import xlsxwriter


def main():
    market_data = get_market_info()
    structured_data = structure_market_data(market_data)
    create_file(structured_data)


def get_market_info():
    # dow = pd.read_csv('^dji.csv', index_col=0)
    
    dji = yf.Ticker('^dji')
    dow = dji.history(period='max')    

    return dow

def structure_market_data(market_data):

    close_volume = market_data[["Close", "Volume"]]
    open_close = market_data[["Open", "Close"]]

    prior1 = close_volume["Close"].shift(periods=1)
    close_volume["prior1"] = prior1
    minus1 = close_volume["Close"] - close_volume["prior1"]
    close_volume["minus1"] = minus1

    prior4 = close_volume["Close"].shift(periods=4)
    close_volume["prior4"] = prior4
    minus4 = close_volume["Close"] - close_volume["prior4"]
    close_volume["minus4"] = minus4

    prior19 = close_volume["Close"].shift(periods=19)
    close_volume["prior19"] = prior19
    minus19 = close_volume["Close"] - close_volume["prior19"]
    close_volume["minus19"] = minus19

    prior63 = close_volume["Close"].shift(periods=63)
    close_volume["prior63"] = prior63
    minus63 = close_volume["Close"] - close_volume["prior63"]
    close_volume["minus63"] = minus63

    prior124 = close_volume["Close"].shift(periods=124)
    close_volume["prior124"] = prior124
    minus124 = close_volume["Close"] - close_volume["prior124"]
    close_volume["minus124"] = minus124

    structured_data = close_volume.drop(columns=['Volume', 'prior1', 'prior4', 'prior19', 'prior63', 'prior124'])

    return structured_data


def create_file(structured_data):

    writer = pd.ExcelWriter('dji_analysis.xlsx', engine='xlsxwriter')
    structured_data.to_excel(writer, sheet_name='dji')
    workbook = writer.book
    worksheet = writer.sheets['dji']
    # worksheet.conditional_format('C:C', {'type' : 'no_blanks'}) # doesn't work

    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})

    # Add a format. Green fill with dark green text.
    format2 = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100'})

    worksheet.conditional_format('C3:G8924', {'type': 'cell','criteria': '<','value': 0,'format': format1})

    # Write another conditional format over the same range.
    worksheet.conditional_format('C3:G8924', {'type': 'cell','criteria': '>','value': 0,'format': format2})

    writer.save()



if __name__ == "__main__":
    main()