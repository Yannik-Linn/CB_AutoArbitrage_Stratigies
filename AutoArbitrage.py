import os  # Miscellaneous operating system interfaces, Lib/os.py
# import random # Generate pseudo-random numbers, Lib/random.py
import time
import datetime  # Basic date and time types, Lib/datetime.py
# Python job scheduling for humans. Run Python functions (or any other callable) periodically using a friendly syntax. https://github.com/dbader/schedule
import schedule
import webbrowser  # Convenient web-browser controller, Lib/webbrowser.py
import xlwings  # xlwings - Make Excel Fly! https://docs.xlwings.org/en/stable/index.html
import pandas
import numpy as np
from scipy.stats import norm
import requests  # HTTP for Humans, https://requests.readthedocs.io/en/latest/
import pysnowball  # snowball's Python API, https://github.com/uname-yang/pysnowball
import browser_cookie3  # Loads cookies used by your web browser into a cookiejar object, https://github.com/borisbabic/browser_cookie3
from chinese_calendar import is_workday # determine workdays in China from 2004 to 2023, https://github.com/LKI/chinese-calendar

from requests.packages.urllib3.exceptions import InsecureRequestWarning # urllib3, HTTP library with thread-safe connection pooling, file post, and more, https://github.com/urllib3/urllib3
requests.packages.urllib3.disable_warnings(InsecureRequestWarning) # Disable any phantom warnings via the PYTHONWARINGS environment variable


source_range_convertible_bond = 'B8:T'   # Get the excel range


# Get the upper limits and weights of the trading strategies from excel
def get_convertible_bond_factor(factor: str):
    factor = factor.split(',', -1)
    return float(factor[0]), float(factor[1])


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

    print("Refresh Convertible Bond Data")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
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
                                          'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Underlying Stock']  # '振幅', '正股代码']
                                 )  # Build up the CB table, with CBs in raws and their data in columns.
    refresh_time = str(time.strftime(
        "%Y%m%d-%H.%M.%S", time.localtime()))  # set the formate of refresh time
    sheet_fund.range('V5').value = 'T_refresh:' + refresh_time  # mark the refresh time in excel
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
        data_fund.loc[i, 'Premium Rate'] = row1["premium_rate"] / 100 if row1["premium_rate"] != None else '停牌' # ‘suspended’
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
        log_str = 'No.' + format(str(i+1), "<6") + format(str(fund_code_str), "<10") \
                  + format(data_fund.loc[i, 'Name'], "<15") \
                  + 'Current: ' + format(str(row1["current"]), "<10") \
                  + 'Premium Rate(%): ' + format(str(row1["premium_rate"]), "<10") \
                  + 'Daily Trend(%): ' + format(str(row1["percent"]), "<10")
        data_fund.loc[i, 'Underlying Stock'] = row1["underlying_symbol"]  # get the underlying stock quote
        print(log_str)   # display the key data in the console: name, current, Premium rate, daily trend
        print(log_str, file=log_file)  # save the log into txt

    data_fund = data_fund.sort_values(by='Premium Rate')  # sort all bonds by Preimum rate, ascending
    data_fund.reset_index(drop=True, inplace=True)
    data_fund.index += 1
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()
    sheet_fund.range('A7').value = data_fund     # update the Excel sheet
    
    data_fund_destination = data_fund[[
        'Underlying Stock', 'Current', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Remain Year']]     
    sheet_dest = wb.sheets['RealTimeData_Stock'] # Save the above selected data into 'RealTimeData_Stock' sheet
    sheet_dest.range('A7').value = data_fund_destination[:30]
    wb.save()    


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
        p = (S*norm.cdf(d1, 0.0, 1.0) - K*np.exp(-r*T)*norm.cdf(d2, 0.0, 1.0))       
    elif option == 'put':
        p = (K*np.exp(-r*T)*norm.cdf(-d2, 0.0, 1.0) - S*np.exp(-q*T)*norm.cdf(-d1, 0.0, 1.0))
    else:
        return None
    return d1, d2, p


@xlwings.func
# Update the real-time data of the underlying stocks & calculate Option Value and bond value
# Rank the convertible bonds by the bias between therotical value and current price
# Including stock price, dividend, volatility, Option value, Pure bond value, Putable price, Callable Price, bias, etc.
def refresh_underlying_stock():
    print("Refresh Underlying Stock Data")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    source_sheets = 'RealTimeData - Stock'
    sheet_stock = wb.sheets[source_sheets]
    source_range = source_range_convertible_bond + str(sheet_stock.used_range.last_cell.row)  # Returns the bottom right cell of the specified range. Read-only.
    print('Data Sheet Range：' + source_range)
    data_stock = pandas.DataFrame(sheet_stock.range(source_range).value,
                                  columns=['Underlying Stock', 'Current', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Remain Year',    # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',   
                                          'Name', 'S_Current', 'Dividend', 'Interest Rate', 'Volatility', 'D1', 'D2',     # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                          'Putable Value', 'Callable Price', 'Pure Bond Value', 'Option Value', 'Theoretical Value', 'Bias']  # '振幅', '正股代码']
                                  )  # Build up the stock table, with stocks in raws and their data in columns.

    for i, stock_code in enumerate(data_stock['Underlying Stock']):
        detail = pandas.DataFrame(pysnowball.quote_detail(stock_code))  # # Get all available data of the bond from the data source
        row2 = detail.loc["quote"][0]
        
        data_stock.loc[i, 'Name'] = row2["name"] # Write the data into data_stock
        data_stock.loc[i, 'S_Current'] = row2["current"]
        data_stock.loc[i, 'Dividend'] = row2["dividend_yield"] 
        data_stock.loc[i, 'D1'], data_stock.loc[i, 'D2'], data_stock.loc[i, 'Option Value'] = bs_option(row2["current"], data_stock.loc[i, 'Conversion Price'],  \      # Calculate the BS option price
                                                                                                        data_stock.loc[i, 'Remain Year'], data_stock.loc[i, 'Interest Rate'] / 100, \
                                                                                                        row2["dividend_yield"] / 100, data_stock.loc[i, 'Volatility'] / 100, option='call')

        log_str = 'No.' + format(str(i+1), "<6") + format(data_stock.loc[i, 'Underlying Stock'], "<15") \
                  + 'CB_Current: ' + format(str(data_stock.loc[i, 'Current']), "<10") \
                  + 'Stock_Current: ' + format(str(row2["current"]), "<10") \
                  + 'V_option: ' + format(str(data_stock.loc[i, 'Option Value']), "<10")  # display the key data in the console: Quote, CB current, Stock current, Option value
        print(log_str)        

    data_stock = data_stock.sort_values(by='Bias')  # sort bt Bias, ascending
    data_stock.reset_index(drop=True, inplace=True)
    data_stock.index += 1
    
    sheet_stock.range('A7').value = data_stock
    wb.save()    


@xlwings.func
# update the CB ranking based on the strategy "multifactor1"
def refresh_multifactor1_convertible_bond():
    print("------Refresh CB: [Multifactor Model 1] Strategy------")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    
    sheet_src = wb.sheets['RealTimeData - Convertible Bond']
    source_range = source_range_convertible_bond + \
        str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                 columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                          'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                          'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Underlying Stock']  # '振幅', '正股代码']
                                 )
    data_fund_destination = data_fund_source[[
        'Quote', 'Name', 'Current', 'Premium Rate', 'Outstanding Amount (m)']]
    threshold_current_price, weight_current_price = get_convertible_bond_factor(
        sheet_src.range('D5').value)                                        # 当前价 阈值，权重： 多因子2 ： D6 多因子1： D5
    threshold_premium_rate, weight_premium_rate = get_convertible_bond_factor(
        sheet_src.range('H5').value)                                       # 溢价 阈值，权重： 多因子2： H6  多因子1： H5
    threshold_outstanding_amt, weight_outstanding_amt = get_convertible_bond_factor(
        sheet_src.range('M5').value)                                       # 剩余规模 阈值，权重： 多因子2： L6  多因子1： L5
    
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < threshold_current_price) &
                                                  (data_fund_destination['Premium Rate'] < threshold_premium_rate) &
                                                  (data_fund_destination['Outstanding Amount (m)'] < threshold_outstanding_amt)]
    data_fund_destination = data_fund_destination.sort_values(
        by='Premium Rate')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)
    
    sheet_dest = wb.sheets['LowPremium_CB_Rotation']     # update the 'LowPremium_CB_Rotation' excel sheet
    sheet_dest.range('Q2').value = data_fund_destination[:20]
    wb.save()


@xlwings.func
# update the CB ranking based on the strategy "multifactor2"
def refresh_multifactor2_convertible_bond():

    print(
        "------Refresh CB: [Multifactor Model 2] Strategy------")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    
    sheet_src = wb.sheets['RealTimeData - Convertible Bond']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                 columns=['Quote', 'Name', 'Current', 'Change', 'Conversion Price', 'Conversion Value', 'Premium Rate', 'Double Low',        # columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                          'Issue Date', 'Maturity Date', 'Remain Year', 'Outstanding Amount (m)', 'Amount (k)', 'Turnover Rate',    # '发行时间', '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                          'Benefit Before Tax', 'Day High', 'Day Low', 'Amplitude', 'Underlying Stock']  # '振幅', '正股代码']
                                 )  
    data_fund_destination = data_fund_source[[
        'Quote', 'Name', 'Current', 'Premium Rate', 'Outstanding Amount (m)']]
    threshold_current_price, weight_current_price = get_convertible_bond_factor(
        sheet_src.range('D6').value)                                        # 当前价阈值，当前价权重： 多因子2 ： D6 多因子1： D5
    threshold_premium_rate, weight_premium_rate = get_convertible_bond_factor(
        sheet_src.range('H6').value)                                        # 溢价阈值，溢价权重： 多因子2： H6  多因子1： H5
    threshold_outstanding_amt, weight_outstanding_amt = get_convertible_bond_factor(
        sheet_src.range('M6').value)                                        # 剩余规模，规模权重： 多因子2： L6  多因子1： L5
    data_fund_destination = data_fund_destination[(data_fund_destination['Current'] < threshold_current_price) &
                                                  (data_fund_destination['Premium Rate'] < threshold_premium_rate) &
                                                  (data_fund_destination['Outstanding Amount (m)'] < threshold_outstanding_amt)]
    data_fund_destination = data_fund_destination.sort_values(by='Premium Rate')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)
    
    sheet_dest = wb.sheets['DoubleLow_CB_Rotation']     # update the 'DoubleLow_CB_Rotation' sheet
    sheet_dest.range('R2').value = data_fund_destination[:20]
    wb.save()



# main function
def main_function():

    # date = datetime.datetime.now().date()    # Chinese Calendar还没更新，所以这三行先不执行
    # if not is_workday(date):
    #     return
    webbrowser.open("https://xueqiu.com/")
    
    for eachfile in os.listdir('./'):
        filename = os.path.join('./', eachfile)
        if os.path.isfile(filename) and filename.startswith("./log"):
            os.remove(filename)     # Delete old log files in the filepath
            
    refresh_convertible_bond()
    refresh_multifactor1_convertible_bond()
    refresh_multifactor2_convertible_bond()   
    refresh_underlying_stock()
    

def main():

    main_function()
    # schedule.every().day.at("07:00").do(main_function)  # 部署7：00执行更新数据任务
    # while True:
    #     schedule.run_pending()


if __name__ == "__main__":
    main()
