import os  # Miscellaneous operating system interfaces, Lib/os.py  ## import random # Generate pseudo-random numbers, Lib/random.py
import time
import datetime  # Basic date and time types, Lib/datetime.py
import schedule # Python job scheduling for humans. Run Python functions (or any other callable) periodically using a friendly syntax. https://github.com/dbader/schedule
import webbrowser  # Convenient web-browser controller, Lib/webbrowser.py
import xlwings  # xlwings - Make Excel Fly! https://docs.xlwings.org/en/stable/index.html
import pandas
import numpy as np
import requests  # HTTP for Humans, https://requests.readthedocs.io/en/latest/
import pysnowball  # snowball's Python API, https://github.com/uname-yang/pysnowball
import browser_cookie3  # Loads cookies used by your web browser into a cookiejar object, https://github.com/borisbabic/browser_cookie3
from scipy.stats import norm
from chinese_calendar import is_workday # determine workdays in China from 2004 to 2024, https://github.com/LKI/chinese-calendar
from requests.packages.urllib3.exceptions import InsecureRequestWarning # urllib3, HTTP library with thread-safe connection pooling, file post, and more, https://github.com/urllib3/urllib3
requests.packages.urllib3.disable_warnings(InsecureRequestWarning) # Disable any phantom warnings via the PYTHONWARINGS environment variable


source_range_convertible_bond = 'B8:T'   # Get the excel range in real time data sheet
source_range_underlyings = 'B8:X'   # Get the excel range in underlyings sheet


@xlwings.func
# Get xq_a_token from data source xueqiu.com
def get_xq_a_token():
    str_xq_a_token = ';'
    while True:
        cj = browser_cookie3.load()
        for item in cj:
            if item.name == "xq_a_token":
                print('get token, %s = %s' % (item.name, item.value))
                str_xq_a_token = 'xq_a_token=' + item.value + ';'
                return str_xq_a_token
        if str_xq_a_token == ";":
            print('get token, retrying ......')
            webbrowser.open("https://xueqiu.com/")
            time.sleep(60)  # Sleep after 60s retries


@xlwings.func
# Update the real-time data of the selected convertible bonds
# Including price, change, remain life, turnover, outstanding amount, Premium rate, benefit before tax, etc.
def refresh_convertible_bond():

    print("------------ Refresh Convertible Bond Data ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    source_sheets = 'RealTimeData_ConvertibleBond'
    sheet_fund = wb.sheets[source_sheets]
    source_range = source_range_convertible_bond + str(sheet_fund.used_range.last_cell.row)  # Returns the bottom right cell of the specified range. Read-only.
    print('Data Sheet Range：' + source_range)
    data_fund = pandas.DataFrame(sheet_fund.range(source_range).value,
                                 columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                          'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                          'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Stock Quote']  # '振幅', '正股代码']
                                 )  # Build up the CB table, with CBs in raws and their data in columns.
    refresh_time = str(time.strftime("%Y%m%d-%H.%M.%S", time.localtime()))  # set the formate of refresh time
    sheet_fund.range('S4').value = 'T_refresh:' + refresh_time  # mark the refresh time in excel
    log_file = open('log_' + source_sheets + '_' + refresh_time +
                    '.txt', 'a+', encoding='utf-8')  # Create a log file "log_RealTimeData_ConvertibleBond_"

    for i, fund_code in enumerate(data_fund['Quote']):
        if str(fund_code).startswith('11') or str(fund_code).startswith('13'):
            fund_code_str = ('SH' + str(fund_code))[0:8]  # SH
        elif str(fund_code).startswith('12'):
            fund_code_str = ('SZ' + str(fund_code))[0:8]  # SZ
        detail = pandas.DataFrame(pysnowball.quote_detail(fund_code_str))  # Get all available data of the bond from the data source
        row1 = detail.loc["quote"][0]
        data_fund.loc[i, 'Name'] = row1["name"] # write the data into data_fund
        data_fund.loc[i, 'Current'] = row1["current"]
        data_fund.loc[i, 'Change'] = row1["percent"] / 100 if row1["percent"] != None else '停牌'  # ‘suspended’
        data_fund.loc[i, 'Conversion Price'] = row1["conversion_price"]
        data_fund.loc[i, 'Conversion Value'] = row1["conversion_value"]
        data_fund.loc[i, 'Premium Rate'] = row1["premium_rate"] if row1["premium_rate"] != None else '停牌' # ‘suspended’
        data_fund.loc[i, 'Double Low'] = row1["current"] + row1["premium_rate"]  
        data_fund.loc[i, 'Issue Date'] = str(time.strftime("%Y-%m-%d", time.localtime(row1["issue_date"]/1000)))
        data_fund.loc[i, 'Maturity Date'] = str(time.strftime("%Y-%m-%d", time.localtime(row1["maturity_date"]/1000)))
        data_fund.loc[i, 'Remain Year'] = row1["remain_year"]
        data_fund.loc[i, 'Outstanding Amount (m)'] = row1["outstanding_amt"] / 1000000 if row1["outstanding_amt"] != None else 1
        data_fund.loc[i, 'Amount (k)'] = row1["amount"] / 1000 if row1["amount"] != None else 0
        data_fund.loc[i, 'Turnover Rate'] = (data_fund.loc[i, 'Amount (k)'] / 1000 / row1["current"]) / (data_fund.loc[i, 'Outstanding Amount (m)'] / 100)
        data_fund.loc[i, 'Benefit Before Tax'] = row1["benefit_before_tax"] / 100 if row1["benefit_before_tax"] != None else '停牌'
        data_fund.loc[i, 'Day High'] = row1["high"]
        data_fund.loc[i, 'Day Low'] = row1["low"]
        if row1["high"] and row1["low"]:
            data_fund.loc[i, 'Amplitude'] = (row1["high"] - row1["low"]) / row1["low"]
        else:
            data_fund.loc[i, 'Amplitude'] = '停牌'
        log_str = format(str(i+1), "<5") + format(str(fund_code_str), "<10") \
                  + format(data_fund.loc[i, 'Name'], "<10") \
                  + 'Current: ' + format(str(row1["current"]), "<10") \
                  + 'Daily Trend(%): ' + format(str(row1["percent"]), "<10") \
                  + 'Premium Rate(%): ' + format(str(row1["premium_rate"]), "<10")
        data_fund.loc[i, 'Stock Quote'] = row1["underlying_symbol"]  # get the underlying stock quote
        print(log_str)   # display the key data in the console: name, current, Premium rate, daily trend
        print(log_str, file=log_file)  # save the log into txt

    data_fund = data_fund.sort_values(by='Premium Rate')  # sort all bonds by Preimum rate, ascending
    data_fund.reset_index(drop=True, inplace=True)
    data_fund.index += 1
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()
    sheet_fund.range('A7').value = data_fund     # update the Excel sheet
    
    data_stock_destination = data_fund[[
        'Quote', 'Name', 'Current', 'Conversion Price', 'Conversion Value', 'Remain Year', 'Premium Rate', 'Stock Quote']]
    data_stock_destination = data_stock_destination.sort_values(by='Quote')  # sort all bonds by Quote, ascending     
    sheet_dest = wb.sheets['Underlying_Values'] # Save the above selected data into 'Underlying_Values' sheet
    sheet_dest.range('A7').value = data_stock_destination[:30]
    wb.save()    


# Calculate Option value based on Black-Scholes model
def bs_option(S, K, T, r, q, sigma, option='call'):
    """
    S: spot price of the underlying stock
    K: strike price, i.e., the conversion price of Convertible Bonds
    T: time to maturity, the remain year of the bonds
    r: risk-free interest rate, here refer to the GCNY10, China 10-Year Government Bond Yield
    q: rate of continuous dividend, of the underlying stock
    sigma: standard deviation of price of underlying asset, based on the historical prices in the last 12 months
    """
    d1 = (np.log(S/K) + (r - q + 0.5*sigma**2)*T)/(sigma*np.sqrt(T))
    d2 = (np.log(S/K) + (r - q - 0.5*sigma**2)*T)/(sigma*np.sqrt(T)) # d2 = d1 - sigma*np.sqrt(T)

    if option == 'call':
        p = (S*np.exp(-q*T)*norm.cdf(d1, 0.0, 1.0) - K*np.exp(-r*T)*norm.cdf(d2, 0.0, 1.0))       
    elif option == 'put':
        p = (K*np.exp(-r*T)*norm.cdf(-d2, 0.0, 1.0) - S*np.exp(-q*T)*norm.cdf(-d1, 0.0, 1.0))
    else:
        return None
    return d1, d2, p

# Calculate implied volatility based on Bisection method
def implied_volatility(P, S, K, T, r, q, option='call'):
    sigma_min = 0.00001
    sigma_max = 1.000
    sigma_mid = (sigma_min + sigma_max) / 2
    
    if option == 'call':
        p_min = bs_option(S, K, T, r, q, sigma_min, option='call')[2]
        p_max = bs_option(S, K, T, r, q, sigma_max, option='call')[2]
        p_mid = bs_option(S, K, T, r, q, sigma_mid, option='call')[2]
        diff = P - p_mid
        
        # if P < p_min or P > p_max:
            # print('Attention, Option Price is beyond the limit, "American Option Case"')
        
        Count = 0
        while abs(diff) > 1e-6:
            if P > p_mid:
                sigma_min = sigma_mid
            else:
                sigma_max = sigma_mid
            sigma_mid = (sigma_min + sigma_max) / 2
            p_mid = bs_option(S, K, T, r, q, sigma_mid, option='call')[2]
            diff = P - p_mid
            Count += 1
            if Count > 100:  
                sigma_mid = 0
                return sigma_mid
    else:
        p_min = bs_option(S, K, T, r, q, sigma_min, option='put')[2]
        p_max = bs_option(S, K, T, r, q, sigma_max, option='put')[2]
        p_mid = bs_option(S, K, T, r, q, sigma_mid, option='put')[2]
        diff = P - p_mid
        
        if P < p_min or P > p_max:
            print('Attention, Option Price is beyond the limit, "American Option Case"')
        
        while abs(diff) > 1e-6:
            if P > p_mid:
                sigma_min = sigma_mid
            else:
                sigma_max = sigma_mid
            sigma_mid = (sigma_min + sigma_max) / 2
            p_mid = bs_option(S, K, T, r, q, sigma_mid, option='put')[2]
            diff = P - p_mid
            Count += 1
            if Count > 100:  
                sigma_mid = 1
                return sigma_mid           
    return sigma_mid

# Calculate Distance to Default based on Merton Model
def Merton_DtD(S,K,T,r,q,sigma):
    """
    S: spot conversion value of the Convertible Bonds
    K: Pure bond value of the Convertible Bonds
    T: time to maturity, the remain year of the bonds
    r: risk-free interest rate, here refer to the GCNY10, China 10-Year Government Bond Yield
    q: rate of continuous dividend, of the underlying stock
    sigma: standard deviation of price of underlying asset, based on the historical prices in the last 12 months
    """
    DtD = (np.log(S/K) + (r - q - 0.5*sigma**2)*T)/(sigma*np.sqrt(T))
    return DtD


@xlwings.func
# Update the real-time data of the underlying stocks & calculate Option Value and bond value
# Rank the convertible bonds by the bias between therotical value and current price
# Including stock price, dividend, volatility, Option value, Pure bond value, Putable price, Callable Price, bias, etc.
def refresh_underlying_values():
    print("------------ Refresh Underlying Values ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    source_sheets = 'Underlying_Values'
    sheet_stock = wb.sheets[source_sheets]
    source_range = source_range_underlyings + str(sheet_stock.used_range.last_cell.row)  # Returns the bottom right cell of the specified range. Read-only.
    print('Data Sheet Range：' + source_range)
    data_stock = pandas.DataFrame(sheet_stock.range(source_range).value,
                                  columns=['Quote', 'Name', 'Current', 'Conversion Price', 'Conversion Value', 'Remain Year', 
                                           'Premium Rate', 'Stock Quote', 'Stock Name', 'Stock Current', 'Dividend', 
                                          'Interest Rate', 'Realized Volatility', 'Implied Volitality', 'Differential Volitality', 
                                          'Putable Price', 'Callable Price', 'Straight Bond Value', 'Option Value', 'Option Price', 
                                          'Theoretical Value', 'Bias', 'DtD']
                                  )  # Build up the stock table, with stocks in raws and their data in columns.
    # data_stock = data_stock.loc[:,['Quote', 'Name', 'Current', 'Stock Quote', 'Stock Name', 'Stock Current',
    #           'Conversion Price', 'Conversion Value', 'Premium Rate', 'Remain Year',
    #         'Dividend', 'Interest Rate', 'Volatility', 'Implied Volitality', 'Diff Volitality',
    #         'Putable Price', 'Callable Price', 'Pure Bond Value', 'Option Value', 'Option Price',
    #         'Theoretical Value', 'Bias', 'DtD']]
    refresh_time = str(time.strftime("%Y%m%d-%H.%M.%S", time.localtime()))  # set the formate of refresh time
    sheet_stock.range('G4').value = 'T_refresh:' + refresh_time  # mark the refresh time in excel
    for i, stock_code in enumerate(data_stock['Stock Quote']):
        detail = pandas.DataFrame(pysnowball.quote_detail(stock_code))  # # Get all available data of the bond from the data source
        row2 = detail.loc["quote"][0]
        
        data_stock.loc[i, 'Stock Name'] = row2["name"] # Write the data into data_stock
        data_stock.loc[i, 'Stock Current'] = row2["current"]
        data_stock.loc[i, 'Dividend'] = row2["dividend_yield"] 
        data_stock.loc[i, 'Option Value'] = 100 / data_stock.loc[i, 'Conversion Price'] * bs_option(row2["current"], data_stock.loc[i, 'Conversion Price'], data_stock.loc[i, 'Remain Year'], data_stock.loc[i,'Interest Rate']/100, row2["dividend_yield"]/100, data_stock.loc[i, 'Realized Volatility']/100, option='call')[2]  # Calculate the BS option price
        data_stock.loc[i, 'Option Price'] = (data_stock.loc[i, 'Current'] - data_stock.loc[i, 'Straight Bond Value']) 
        data_stock.loc[i, 'Implied Volitality'] = 100*implied_volatility(data_stock.loc[i, 'Option Price'] * data_stock.loc[i, 'Conversion Price'] / 100, row2["current"], data_stock.loc[i, 'Conversion Price'], \
                                                                     data_stock.loc[i, 'Remain Year'], data_stock.loc[i, 'Interest Rate'] / 100, \
                                                                     row2["dividend_yield"] / 100, option='call')
        data_stock.loc[i, 'Differential Volitality'] = (data_stock.loc[i, 'Implied Volitality'] - data_stock.loc[i, 'Realized Volatility']) / 100
        data_stock.loc[i, 'DtD'] = Merton_DtD(data_stock.loc[i, 'Conversion Value'],data_stock.loc[i, 'Straight Bond Value'],  \
                                              data_stock.loc[i, 'Remain Year'], data_stock.loc[i, 'Interest Rate'] / 100, \
                                              row2["dividend_yield"] / 100, data_stock.loc[i, 'Realized Volatility'] / 100)
        data_stock.loc[i, 'Theoretical Value'] = data_stock.loc[i, 'Straight Bond Value'] + data_stock.loc[i, 'Option Value']
        data_stock.loc[i, 'Bias'] = data_stock.loc[i, 'Current'] / data_stock.loc[i, 'Theoretical Value'] - 1
        log_str = format(str(i+1), "<5") + format(data_stock.loc[i, 'Name'], "<10") \
              + 'Option Value: ' + format(data_stock.loc[i, 'Option Value'], '<10.2f') \
              + 'Bond Value: ' + format(data_stock.loc[i, 'Straight Bond Value'], "<10.2f") \
              + 'Diff Vol.: ' + format(data_stock.loc[i, 'Differential Volitality'], "<10.2f") \
              + 'DtD: ' + format(data_stock.loc[i, "DtD"], "<10.2f") \
              + 'Bias: ' + format(data_stock.loc[i, 'Bias'], "<10.2%")  # display the key data in the console: Quote, CB current, Stock current, Option value
        print(log_str)        

    data_stock = data_stock.sort_values(by='Quote')  # sort bt Quote, ascending

    data_stock.reset_index(drop=True, inplace=True)
    data_stock.index += 1
    
    sheet_stock.range('A7').value = data_stock
    wb.save()    


@xlwings.func
# Refresh CB ranking based on [Low Premium Rate] Strategy
def refresh_premium_rate():
    print("------------ Refresh Ranking: [Low Premium Rate] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['RealTimeData_ConvertibleBond']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                                'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                                'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Stock Quote'])
    data_fund_destination = data_fund_source[['Quote','Name','Current','Premium Rate','Outstanding Amount (m)']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < sheet_src.range('D2').value) &
                                                  (data_fund_destination['Premium Rate'] < sheet_src.range('H2').value) &
                                                  (data_fund_destination['Outstanding Amount (m)'] < sheet_src.range('M2').value)]
    data_fund_destination = data_fund_destination.sort_values(by='Premium Rate')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)

    sheet_dest = wb.sheets['Singlefactor Strategies']     # Update Excel sheet: 'Singlefactor Strategies'
    sheet_dest.range('J2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# Refresh CB ranking based on [Low Current Price + Low Premium Rate * 100] Strategy
def refresh_DoubleLow():
    print("------------ Refresh Ranking: [Double Low] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['RealTimeData_ConvertibleBond']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                                'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                                'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Stock Quote'])
    data_fund_destination = data_fund_source[['Quote','Name','Current','Premium Rate','Double Low','Outstanding Amount (m)']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < sheet_src.range('D3').value) &
                                                  (data_fund_destination['Premium Rate'] < sheet_src.range('H3').value) &
                                                  (data_fund_destination['Outstanding Amount (m)'] < sheet_src.range('M3').value)]
    data_fund_destination = data_fund_destination[['Quote','Name','Current','Double Low','Outstanding Amount (m)']]
    data_fund_destination = data_fund_destination.sort_values(by='Double Low')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    sheet_dest = wb.sheets['Singlefactor Strategies']     # Update Excel sheet: 'Singlefactor Strategies'
    sheet_dest.range('R2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# Refresh CB ranking based on [Highest Differential Volatility] Strategy
def refresh_diff_volatility():
    print("------------ Refresh Ranking: [High Differential Volatility] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['Underlying_Values']
    source_range = source_range_underlyings + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['Quote', 'Name', 'Current', 'Conversion Price', 'Conversion Value', 'Remain Year', 
                                                'Premium Rate', 'Stock Quote', 'Stock Name', 'Stock Current', 'Dividend', 
                                               'Interest Rate', 'Realized Volatility', 'Implied Volitality', 'Differential Volitality', 
                                               'Putable Price', 'Callable Price', 'Straight Bond Value', 'Option Value', 'Option Price', 
                                               'Theoretical Value', 'Bias', 'DtD'])
    data_fund_destination = data_fund_source[['Quote','Name','Current','Realized Volatility','Implied Volitality','Differential Volitality']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < sheet_src.range('D2').value) &
                                                  (data_fund_destination['Implied Volitality'] > 0) &
                                                  (data_fund_destination['Differential Volitality'] < sheet_src.range('P2').value)]
    data_fund_destination = data_fund_destination.sort_values(by='Differential Volitality',ascending=True)
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    sheet_dest = wb.sheets['Singlefactor Strategies']     # Update Excel sheet: 'Singlefactor Strategies'
    sheet_dest.range('A20').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# Refresh CB ranking based on [Lowest Current Price / Theortical Value - 1 ] Strategy
def refresh_Bias():
    print("------------ Refresh Ranking: [Low Price/Value Bias] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['Underlying_Values']
    source_range = source_range_underlyings + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['Quote', 'Name', 'Current', 'Conversion Price', 'Conversion Value', 'Remain Year', 
                                                'Premium Rate', 'Stock Quote', 'Stock Name', 'Stock Current', 'Dividend', 
                                               'Interest Rate', 'Realized Volatility', 'Implied Volitality', 'Differential Volitality', 
                                               'Putable Price', 'Callable Price', 'Straight Bond Value', 'Option Value', 'Option Price', 
                                               'Theoretical Value', 'Bias', 'DtD'])
    data_fund_destination = data_fund_source[['Quote','Name','Current','Theoretical Value','Bias']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < sheet_src.range('D3').value) &
                                                  (data_fund_destination['Bias'] < sheet_src.range('W3').value)]
    data_fund_destination = data_fund_destination.sort_values(by='Bias')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    sheet_dest = wb.sheets['Singlefactor Strategies']     # Update Excel sheet: 'Singlefactor Strategies'
    sheet_dest.range('J20').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# Refresh CB ranking based on [Highest Distace to Default DtD] Strategy
def refresh_DtD():
    print("------------ Refresh Ranking: [Hight Distace to Default DtD] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['Underlying_Values']
    source_range = source_range_underlyings + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['Quote', 'Name', 'Current', 'Conversion Price', 'Conversion Value', 'Remain Year', 
                                                'Premium Rate', 'Stock Quote', 'Stock Name', 'Stock Current', 'Dividend', 
                                               'Interest Rate', 'Realized Volatility', 'Implied Volitality', 'Differential Volitality', 
                                               'Putable Price', 'Callable Price', 'Straight Bond Value', 'Option Value', 'Option Price', 
                                               'Theoretical Value', 'Bias', 'DtD'])
    data_fund_destination = data_fund_source[['Quote','Name','Current','Straight Bond Value','DtD']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < sheet_src.range('D4').value) &
                                                  (data_fund_destination['Straight Bond Value'] < sheet_src.range('S4').value) &
                                                  (data_fund_destination['DtD'] > sheet_src.range('X4').value)]
    data_fund_destination = data_fund_destination.sort_values(by='DtD',ascending=False)
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    sheet_dest = wb.sheets['Singlefactor Strategies']     # Update Excel sheet: 'Singlefactor Strategies'
    sheet_dest.range('R20').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# General Button to refresh all single factor strategies in the sheet
def refresh_singlefactor_strategies():
    
    refresh_premium_rate()
    refresh_DoubleLow()
    refresh_diff_volatility()
    refresh_Bias()
    refresh_DtD()

# Get the upper limits and weights of the trading strategies from excel
def get_convertible_bond_factor(factor: str):
    factor = factor.split('&', -1)
    return float(factor[0]), float(factor[1])

@xlwings.func
# update the CB ranking based on the strategy "multifactor1"
def refresh_multifactor1_convertible_bond():
    print("------------ Refresh Ranking: [Multifactor Model 1] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    
    sheet_src = wb.sheets['RealTimeData_ConvertibleBond']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                 columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                          'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                          'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Stock Quote']  # '振幅', '正股代码']
                                 )
    threshold_current_price, weight_current_price = get_convertible_bond_factor(sheet_src.range('D5').value) # get the threhold and weight parameters
    threshold_premium_rate, weight_premium_rate = get_convertible_bond_factor(sheet_src.range('H5').value)
    threshold_outstanding_amt, weight_outstanding_amt = get_convertible_bond_factor(sheet_src.range('M5').value)

    for i, fund_code in enumerate(data_fund_source['Quote']):
        data_fund_source.loc[i, 'multifactor1'] = data_fund_source.loc[i, 'Current'] * weight_current_price
        + data_fund_source.loc[i, 'Premium Rate'] * weight_premium_rate
        + data_fund_source.loc[i, 'Outstanding Amount (m)'] * weight_outstanding_amt
    data_fund_destination = data_fund_source[['Quote', 'Name', 'Current', 'Premium Rate', 'Outstanding Amount (m)', 'multifactor1']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < threshold_current_price) &
                                                  (data_fund_destination['Premium Rate'] < threshold_premium_rate) &
                                                  (data_fund_destination['Outstanding Amount (m)'] < threshold_outstanding_amt)]
    data_fund_destination = data_fund_destination.sort_values(by='multifactor1')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)
    
    sheet_dest = wb.sheets['Multifactor Strategies']     # update the 'LowPremium_CB_Rotation' excel sheet
    sheet_dest.range('H2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# update the CB ranking based on the strategy "multifactor2"
def refresh_multifactor2_convertible_bond():
    print("------------ Refresh Ranking: [Multifactor Model 2] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    
    sheet_src = wb.sheets['RealTimeData_ConvertibleBond']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                 columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                          'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                          'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Stock Quote']  # '振幅', '正股代码']
                                 )
    threshold_current_price, weight_current_price = get_convertible_bond_factor(sheet_src.range('D6').value) # get the threhold and weight parameters
    threshold_premium_rate, weight_premium_rate = get_convertible_bond_factor(sheet_src.range('H6').value)
    threshold_outstanding_amt, weight_outstanding_amt = get_convertible_bond_factor(sheet_src.range('M6').value)

    for i, fund_code in enumerate(data_fund_source['Quote']):
        data_fund_source.loc[i, 'multifactor2'] = data_fund_source.loc[i, 'Current'] * weight_current_price
        + data_fund_source.loc[i, 'Premium Rate'] * weight_premium_rate
        + data_fund_source.loc[i, 'Outstanding Amount (m)'] * weight_outstanding_amt
    data_fund_destination = data_fund_source[['Quote', 'Name', 'Current', 'Premium Rate', 'Outstanding Amount (m)', 'multifactor2']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < threshold_current_price) &
                                                  (data_fund_destination['Premium Rate'] < threshold_premium_rate) &
                                                  (data_fund_destination['Outstanding Amount (m)'] < threshold_outstanding_amt)]
    data_fund_destination = data_fund_destination.sort_values(by='multifactor2')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)
    
    sheet_dest = wb.sheets['Multifactor Strategies']     # update the 'LowPremium_CB_Rotation' excel sheet
    sheet_dest.range('Q2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# update the CB ranking based on the strategy "multifactor3"
def refresh_multifactor3_convertible_bond():
    print("------------ Refresh Ranking: [Multifactor Model 3] Strategy ------------")
    xlwings.Book("AutoArbitrage.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    
    sheet_src = wb.sheets['Underlying_Values']
    source_range = source_range_underlyings + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                 columns=['Quote', 'Name', 'Current', 'Conversion Price', 'Conversion Value', 'Remain Year', 
                                          'Premium Rate', 'Stock Quote', 'Stock Name', 'Stock Current', 'Dividend', 
                                         'Interest Rate', 'Realized Volatility', 'Implied Volitality', 'Differential Volitality', 
                                         'Putable Price', 'Callable Price', 'Straight Bond Value', 'Option Value', 'Option Price', 
                                         'Theoretical Value', 'Bias', 'DtD']
                                 )
    threshold_Differential_Volitality, weight_Differential_Volitality = get_convertible_bond_factor(sheet_src.range('P6').value) # get the threhold and weight parameters
    threshold_Bias, weight_Bias = get_convertible_bond_factor(sheet_src.range('W6').value)
    threshold_DtD, weight_DtD = get_convertible_bond_factor(sheet_src.range('X6').value)

    for i, fund_code in enumerate(data_fund_source['Quote']):
        data_fund_source.loc[i, 'multifactor3'] = data_fund_source.loc[i, 'Differential Volitality'] * weight_Differential_Volitality
        + data_fund_source.loc[i, 'Bias'] * weight_Bias
        - data_fund_source.loc[i, 'DtD'] * weight_DtD
    data_fund_destination = data_fund_source[['Quote', 'Name', 'Differential Volitality', 'Bias', 'DtD', 'multifactor3']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Differential Volitality'] < threshold_Differential_Volitality) &
                                                  (data_fund_destination['Bias'] < threshold_Bias) &
                                                  (data_fund_destination['DtD'] > threshold_DtD)]
    data_fund_destination = data_fund_destination.sort_values(by='multifactor3')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)
    
    sheet_dest = wb.sheets['Multifactor Strategies']     # update the 'LowPremium_CB_Rotation' excel sheet
    sheet_dest.range('Z2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# General Button to refresh all multiple-factor strategies in the sheet
def refresh_multifactor_strategies():
    
    refresh_multifactor1_convertible_bond()
    refresh_multifactor2_convertible_bond()   
    refresh_multifactor3_convertible_bond()  

# main function
def main_function():

    # date = datetime.datetime.now().date()    # Holidays suspension
    # if not is_workday(date):
    #     return
    webbrowser.open("https://xueqiu.com/")
    
    for eachfile in os.listdir('./'):
        filename = os.path.join('./', eachfile)
        if os.path.isfile(filename) and filename.startswith("./log"):
            os.remove(filename)     # Delete old log files in the filepath
            
    refresh_convertible_bond()
    refresh_underlying_values()
    
    refresh_premium_rate()
    refresh_DoubleLow()
    refresh_diff_volatility()
    refresh_Bias()
    refresh_DtD()
    refresh_multifactor1_convertible_bond()
    refresh_multifactor2_convertible_bond()   
    refresh_multifactor3_convertible_bond()  
    
def main():

    main_function()
    # schedule.every().day.at("07:00").do(main_function)  # deploy the refresh tasks at 7：00 a.m. everyday
    # while True:
    #     schedule.run_pending()

if __name__ == "__main__":
    main()
