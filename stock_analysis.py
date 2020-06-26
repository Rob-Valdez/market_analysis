# stock_analysis.py

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf
import xlsxwriter


def main():
    market_data = get_market_info()
    structured_data, row_count = structure_market_data(market_data)
    create_file(structured_data, row_count)


def get_market_info():
    
    symbol = input('Enter the ticker for the stock you would like to analyze:  ')
    tick = yf.Ticker(f'{symbol}')
    tick = tick.history(period='max')    

    return tick

def structure_market_data(market_data):

    close_volume = market_data[["Close", "Volume"]]
    open_close = market_data[["Open", "Close"]]

    prior1 = close_volume["Close"].shift(periods=1)
    close_volume["prior1"] = prior1
    minus1 = close_volume["Close"] - close_volume["prior1"]
    close_volume["Delta1D"] = minus1

    prior4 = close_volume["Close"].shift(periods=4)
    close_volume["prior4"] = prior4
    minus4 = close_volume["Close"] - close_volume["prior4"]
    close_volume["Delta1W"] = minus4

    prior19 = close_volume["Close"].shift(periods=19)
    close_volume["prior19"] = prior19
    minus19 = close_volume["Close"] - close_volume["prior19"]
    close_volume["Delta1M"] = minus19

    prior63 = close_volume["Close"].shift(periods=63)
    close_volume["prior63"] = prior63
    minus63 = close_volume["Close"] - close_volume["prior63"]
    close_volume["Delta3M"] = minus63

    prior124 = close_volume["Close"].shift(periods=124)
    close_volume["prior124"] = prior124
    minus124 = close_volume["Close"] - close_volume["prior124"]
    close_volume["Delta6M"] = minus124

    structured_data = close_volume.drop(columns=['Volume', 'prior1', 'prior4', 'prior19', 'prior63', 'prior124'])
    row_count = len(structured_data.index)

    return structured_data, row_count


def create_file(structured_data, row_count):

    row_count += 1
    writer = pd.ExcelWriter('analysis.xlsx', engine='xlsxwriter')
    structured_data.to_excel(writer, sheet_name='analysis')
    workbook = writer.book
    worksheet = writer.sheets['analysis']
    # worksheet.conditional_format('C:C', {'type' : 'no_blanks'}) # doesn't work

    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})

    # Add a format. Green fill with dark green text.
    format2 = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100'})

    worksheet.conditional_format(f'C2:G{row_count}', {'type': 'cell','criteria': '<','value': 0,'format': format1})

    # Write another conditional format over the same range.
    worksheet.conditional_format(f'C2:G{row_count}', {'type': 'cell','criteria': '>','value': 0,'format': format2})

    writer.save()



if __name__ == "__main__":
    main()