import os # Miscellaneous operating system interfaces, Lib/os.py
# import random # Generate pseudo-random numbers, Lib/random.py
import time # Necessary??
import datetime # Basic date and time types, Lib/datetime.py
import schedule # Python job scheduling for humans. Run Python functions (or any other callable) periodically using a friendly syntax. https://github.com/dbader/schedule
import webbrowser # Convenient web-browser controller, Lib/webbrowser.py
import xlwings # xlwings - Make Excel Fly! https://docs.xlwings.org/en/stable/index.html
import pandas
import requests # HTTP for Humans, https://requests.readthedocs.io/en/latest/
import pysnowball # snowball's Python API, https://github.com/uname-yang/pysnowball
import browser_cookie3 # Loads cookies used by your web browser into a cookiejar object, https://github.com/borisbabic/browser_cookie3
from chinese_calendar import is_workday # determine workdays in China from 2004 to 2023, https://github.com/LKI/chinese-calendar

# Disable any phantom warnings via the PYTHONWARINGS ebvironment variable
from requests.packages.urllib3.exceptions import InsecureRequestWarning  # urllib3, HTTP library with thread-safe connection pooling, file post, and more, https://github.com/urllib3/urllib3
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)



@xlwings.func
# Get xq_a_token from xueqiu.com
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
            time.sleep(60) # Sleep after 60s retries


source_range_convertible_bond = 'B8:T'   # 'B8:U'


# 获取可转债实时数据列表上方的多因子策略的阈值上限和权重
def get_convertible_bond_factor(factor:str):
    factor = factor.split(',', -1)
    return float(factor[0]), float(factor[1])


@xlwings.func           # OK
# 更新可转债实时数据：价格、涨跌幅、转股价、转股价值、溢价率、双低值、到期时间、剩余年限、剩余规模、成交金额、换手率、税前收益、振幅等
def refresh_convertible_bond():
    print("Refresh Convertible Bond DATA：价格、涨跌幅、转股价、转股价值、溢价率、双低值、到期时间、剩余年限、剩余规模、成交金额、换手率、税前收益、振幅等")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    source_sheets = '可转债实时数据'
    sheet_fund = wb.sheets[source_sheets]
    source_range = source_range_convertible_bond + \
        str(sheet_fund.used_range.last_cell.row) # Returns the bottom right cell of the specified range. Read-only.
    print('Data Sheet Range：' + source_range)
    data_fund = pandas.DataFrame(sheet_fund.range(source_range).value,
                                  columns=['转债代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                           '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                           '振幅', '多因子1排名', '多因子2排名']
                                 ) # Build up the CB table, with CBs in raws and their data in columns.
    refresh_time = str(time.strftime("%Y%m%d-%H.%M.%S", time.localtime()))  # 设置时间格式
    sheet_fund.range('V5').value = '更新时间:' + refresh_time  # 在表中标记时间
    log_file = open('log_' + source_sheets + '_' + refresh_time +
                    '.txt', 'a+', encoding='utf-8') # Create a log file "log_可转债实时数据_"

    for i, fund_code in enumerate(data_fund['转债代码']):
        if str(fund_code).startswith('11') or str(fund_code).startswith('13'):
            fund_code_str = ('SH' + str(fund_code))[0:8] # SH
        elif str(fund_code).startswith('12'):
            fund_code_str = ('SZ' + str(fund_code))[0:8] # SZ
        detail = pandas.DataFrame(pysnowball.quote_detail(fund_code_str))  # 获取某支股票的行情数据-详细
        row1 = detail.loc["quote"][0]                                      # 获取该股票全部信息； 对拟新增功能，需考察： underlying_symbol， remain_year， pure_bond_price， premium_rate 等等
        data_fund.loc[i, '当前价'] = row1["current"]                        # 将所选信息更新写入data_fund
        data_fund.loc[i, '涨跌幅'] = row1["percent"] / \
            100 if row1["percent"] != None else '停牌' # ‘suspended’
        data_fund.loc[i, '转股价'] = row1["conversion_price"]
        data_fund.loc[i, '转股价值'] = row1["conversion_value"]
        data_fund.loc[i, '溢价率'] = row1["premium_rate"] / \
            100 if row1["premium_rate"] != None else '停牌'
        data_fund.loc[i, '双低值'] = row1["current"] + row1["premium_rate"]
        data_fund.loc[i, '到期时间'] = str(time.strftime(
            "%Y-%m-%d", time.localtime(row1["maturity_date"]/1000)))
        data_fund.loc[i, '剩余年限'] = row1["remain_year"]
        data_fund.loc[i, '剩余规模'] = row1["outstanding_amt"] / \
            100000000 if row1["outstanding_amt"] != None else 1
        data_fund.loc[i, '成交金额'] = row1["amount"] / \
            10000 if row1["amount"] != None else 0
        data_fund.loc[i, '换手率'] = (
            data_fund.loc[i, '成交金额'] / 10000 / row1["current"]) / (data_fund.loc[i, '剩余规模'] / 100)
        data_fund.loc[i, '税前收益'] = row1["benefit_before_tax"] / \
            100 if row1["benefit_before_tax"] != None else '停牌'
        data_fund.loc[i, '最高价'] = row1["high"]
        data_fund.loc[i, '最低价'] = row1["low"]
        if row1["high"] and row1["low"]:
            data_fund.loc[i, '振幅'] = (row1["high"] - row1["low"]) / row1["low"]
        else:
            data_fund.loc[i, '振幅'] = '停牌'
        log_str = 'No.' + format(str(i+1), "<6") + format(str(fund_code_str), "<10") \
                  + format(data_fund.loc[i, '转债名称'], "<15") \
                  + 'Spot Price: ' + format(str(row1["current"]), "<10") \
                  + 'Premium Rate(%): ' + format(str(row1["premium_rate"]), "<10") \
                  + 'Daily Trend(%): ' + format(str(row1["percent"]), "<10")            # 在命令行显示转债名，当前价，溢价率，涨跌幅
        # data_fund.loc[i, '正股代码'] = row1["underlying_symbol"]
        print(log_str)
        print(log_str, file=log_file) # 存入log文档

    data_fund = data_fund.sort_values(by='溢价率') # 以溢价率重新排序
    data_fund.reset_index(drop=True, inplace=True)
    data_fund.index += 1
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()
    # 更新原Excel
    sheet_fund.range('A7').value = data_fund
    wb.save()


@xlwings.func           # OK
# 更新可转债实时数据：价格、涨跌幅、转股价、转股价值、溢价率、双低值、到期时间、剩余年限、剩余规模、成交金额、换手率、税前收益、振幅等
def refresh_optionBS_purebond_value_convertible_bond():
    print("Refresh Underlying Stock DATA：价格、涨跌幅、转股价、转股价值、溢价率、双低值、到期时间、剩余年限、剩余规模、成交金额、换手率、税前收益、振幅等")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    source_sheets = '正股实时数据'
    sheet_fund = wb.sheets[source_sheets]
    source_range = source_range_convertible_bond + \
        str(sheet_fund.used_range.last_cell.row) # Returns the bottom right cell of the specified range. Read-only.
    print('Data Sheet Range：' + source_range)
    data_fund = pandas.DataFrame(sheet_fund.range(source_range).value,
                                  columns=['正股代码', '转债名称', '当前价', '涨跌幅', '转股价', '转股价值', '溢价率', '双低值',
                                           '到期时间', '剩余年限', '剩余规模', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                           '振幅', '多因子1排名', '多因子2排名']
                                 ) # Build up the CB table, with CBs in raws and their data in columns.
    refresh_time = str(time.strftime("%Y%m%d-%H.%M.%S", time.localtime()))  # 设置时间格式
    sheet_fund.range('V5').value = '更新时间:' + refresh_time  # 在表中标记时间
    log_file = open('log_' + source_sheets + '_' + refresh_time +
                    '.txt', 'a+', encoding='utf-8') # Create a log file "log_可转债实时数据_"

    for i, fund_code in enumerate(data_fund['正股代码']):
        # if str(fund_code).startswith('11') or str(fund_code).startswith('13'):
        #     fund_code_str = ('SH' + str(fund_code))[0:8] # SH
        # elif str(fund_code).startswith('12'):
        #     fund_code_str = ('SZ' + str(fund_code))[0:8] # SZ
        detail = pandas.DataFrame(pysnowball.quote_detail(fund_code))  # 获取某支股票的行情数据-详细
        row1 = detail.loc["quote"][0]                                      # 获取该股票全部信息； 对拟新增功能，需考察： underlying_symbol， remain_year， pure_bond_price， premium_rate 等等
        data_fund.loc[i, '当前价'] = row1["current"]                        # 将所选信息更新写入data_fund
        data_fund.loc[i, '涨跌幅'] = row1["percent"] / \
            100 if row1["percent"] != None else '停牌' # ‘suspended’
        data_fund.loc[i, '转股价'] = row1["conversion_price"]
        data_fund.loc[i, '转股价值'] = row1["conversion_value"]
        data_fund.loc[i, '溢价率'] = row1["premium_rate"] / \
            100 if row1["premium_rate"] != None else '停牌'
        data_fund.loc[i, '双低值'] = row1["current"] + row1["premium_rate"]
        data_fund.loc[i, '到期时间'] = str(time.strftime(
            "%Y-%m-%d", time.localtime(row1["maturity_date"]/1000)))
        data_fund.loc[i, '剩余年限'] = row1["remain_year"]
        data_fund.loc[i, '剩余规模'] = row1["outstanding_amt"] / \
            100000000 if row1["outstanding_amt"] != None else 1
        data_fund.loc[i, '成交金额'] = row1["amount"] / \
            10000 if row1["amount"] != None else 0
        data_fund.loc[i, '换手率'] = (
            data_fund.loc[i, '成交金额'] / 10000 / row1["current"]) / (data_fund.loc[i, '剩余规模'] / 100)
        data_fund.loc[i, '税前收益'] = row1["benefit_before_tax"] / \
            100 if row1["benefit_before_tax"] != None else '停牌'
        data_fund.loc[i, '最高价'] = row1["high"]
        data_fund.loc[i, '最低价'] = row1["low"]
        if row1["high"] and row1["low"]:
            data_fund.loc[i, '振幅'] = (row1["high"] - row1["low"]) / row1["low"]
        else:
            data_fund.loc[i, '振幅'] = '停牌'
        log_str = 'No.' + format(str(i+1), "<6") + format(str(fund_code_str), "<10") \
                  + format(data_fund.loc[i, '转债名称'], "<15") \
                  + 'Spot Price: ' + format(str(row1["current"]), "<10") \
                  + 'Premium Rate(%): ' + format(str(row1["premium_rate"]), "<10") \
                  + 'Daily Trend(%): ' + format(str(row1["percent"]), "<10")            # 在命令行显示转债名，当前价，溢价率，涨跌幅
        data_fund.loc[i, '正股代码'] = row1["underlying_symbol"]
        print(log_str)
        print(log_str, file=log_file) # 存入log文档

    data_fund = data_fund.sort_values(by='溢价率') # 以溢价率重新排序
    data_fund.reset_index(drop=True, inplace=True)
    data_fund.index += 1
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()
    # 更新原Excel
    sheet_fund.range('A7').value = data_fund
    wb.save()






@xlwings.func
# 更新低溢价可转债数据
def refresh_premium_rate_convertible_bond():
    print("--------------------------------------Refresh Data: [Low Premium Rate] Convertible Bond----------------------------------------------------")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['可转债实时数据']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['转债代码','转债名称','Spot Price','涨跌幅','转股价','转股价值','Premium Rate','双低值',
                                                '到期时间','剩余年限','Outstanding Amount','成交金额','换手率','税前收益','最高价','最低价',
                                                '振幅','多因子1排名','多因子2排名'])
    data_fund_destination = data_fund_source[['转债代码','转债名称','Spot Price','Premium Rate','Outstanding Amount']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Spot Price'] < sheet_src.range('D2').value) &
                                                  (data_fund_destination['Premium Rate'] < sheet_src.range('H2').value) &
                                                  (data_fund_destination['Outstanding Amount'] < sheet_src.range('L2').value)]
    data_fund_destination = data_fund_destination.sort_values(by='Premium Rate')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)

    # 更新低溢价可转债轮动sheet
    sheet_dest = wb.sheets['低溢价可转债轮动']
    sheet_dest.range('H2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# 更新双低可转债数据
def refresh_price_and_premium_rate_convertible_bond():
    print("--------------------------------------Refresh Data: [Low Premium Rate & Low Spot Price] Convertible Bond----------------------------------")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['可转债实时数据']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['转债代码','转债名称','Spot Price','涨跌幅','转股价','转股价值','Premium Rate','Sum Factor',
                                                '到期时间','剩余年限','Outstanding Amount','成交金额','换手率','税前收益','最高价','最低价',
                                                '振幅','多因子1排名','多因子2排名'])
    data_fund_destination = data_fund_source[['转债代码','转债名称','Spot Price','Premium Rate','Sum Factor','Outstanding Amount']]
    data_fund_destination = data_fund_destination[(data_fund_destination['Spot Price'] < sheet_src.range('D3').value) &
                                                  (data_fund_destination['Premium Rate'] < sheet_src.range('H3').value) &
                                                  (data_fund_destination['Outstanding Amount'] < sheet_src.range('L3').value)]
    data_fund_destination = data_fund_destination.sort_values(by='Sum Factor')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    # 更新双低可转债轮动sheet
    sheet_dest = wb.sheets['双低可转债轮动']
    sheet_dest.range('H2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func
# 更新多因子1可转债数据
def refresh_multifactor1_convertible_bond():
    print("---------------------------------------Refresh Data: [Multifactor Model 1] Convertible Bond------------------------------------------------")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['可转债实时数据']
    source_range = source_range_convertible_bond + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['转债代码','转债名称','Spot Price','涨跌幅','转股价','转股价值','Premium Rate','双低值',
                                                '到期时间','剩余年限','Outstanding Amount','成交金额','换手率','税前收益','最高价','最低价',
                                                '振幅','多因子1排名','多因子2排名'])
    data_fund_destination = data_fund_source[['转债代码','转债名称','Spot Price','Premium Rate','Outstanding Amount']]
    threshold_current_price, weight_current_price = get_convertible_bond_factor(sheet_src.range('D5').value)
    threshold_premium_rate, weight_premium_rate = get_convertible_bond_factor(sheet_src.range('H5').value)
    threshold_outstanding_amt, weight_outstanding_amt = get_convertible_bond_factor(sheet_src.range('L5').value)

    # for i, fund_code in enumerate(data_fund_source['转债代码']):
    #     data_fund_source.loc[i, '多因子1排名'] = data_fund_source.loc[i, '当前价'] * weight_current_price
    #     + data_fund_source.loc[i, '溢价率'] * weight_premium_rate
    #     + data_fund_source.loc[i, '剩余规模'] * weight_outstanding_amt

    data_fund_destination = data_fund_destination[(data_fund_destination['Spot Price'] < threshold_current_price) &
                                                  (data_fund_destination['Premium Rate'] < threshold_premium_rate) &
                                                  (data_fund_destination['Outstanding Amount'] < threshold_outstanding_amt)]
    data_fund_destination = data_fund_destination.sort_values(by='Premium Rate')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)

    # 更新低溢价可转债轮动sheet
    sheet_dest = wb.sheets['低溢价可转债轮动']
    sheet_dest.range('Q2').value = data_fund_destination[:20]
    wb.save()

@xlwings.func           # OK
# 更新多因子2可转债数据
def refresh_multifactor2_convertible_bond():

    print("--------------------------------------Refresh Data: [Multifactor Model 2] Convertible Bond-------------------------------------------------")
    xlwings.Book("UniversalRotation_EN.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['可转债实时数据']
    source_range = source_range_convertible_bond + \
        str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                        columns=['转债代码', '转债名称', 'Spot Price', '涨跌幅', '转股价', '转股价值', 'Premium Rate', '双低值',
                                                  '到期时间', '剩余年限', 'Outstanding Amount', '成交金额', '换手率', '税前收益', '最高价', '最低价',
                                                  '振幅', '多因子1排名', '多因子2排名'])
    data_fund_destination = data_fund_source[[
        '转债代码', '转债名称', 'Spot Price','Premium Rate','Outstanding Amount']]
    threshold_current_price, weight_current_price = get_convertible_bond_factor(
        sheet_src.range('D6').value)                                        # 当前价阈值，当前价权重： 多因子2 ： D6 多因子1： D5
    threshold_premium_rate, weight_premium_rate = get_convertible_bond_factor(
        sheet_src.range('H6').value)                                        # 溢价阈值，溢价权重： 多因子2： H6  多因子1： H5
    threshold_outstanding_amt, weight_outstanding_amt = get_convertible_bond_factor(
        sheet_src.range('L6').value)                                        # 剩余规模，规模权重： 多因子2： L6  多因子1： L5
    data_fund_destination = data_fund_destination[(data_fund_destination['Spot Price'] < threshold_current_price) &
                                                  (data_fund_destination['Premium Rate'] < threshold_premium_rate) &
                                                  (data_fund_destination['Outstanding Amount'] < threshold_outstanding_amt)]  # 这一行问题严重！！！
    data_fund_destination = data_fund_destination.sort_values(by='Premium Rate')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index += 1
    print(data_fund_destination)

    # 更新双低可转债轮动sheet
    sheet_dest = wb.sheets['双低可转债轮动']
    sheet_dest.range('R2').value = data_fund_destination[:20]  # 写入'双低可转债轮动'表的 R2 之后 ？？？
    wb.save()


def main_function():
    
    date = datetime.datetime.now().date()
    if not is_workday(date):
        return
    webbrowser.open("https://xueqiu.com/")

    # Delete old log files in the filepath
    for eachfile in os.listdir('./'):
        filename = os.path.join('./', eachfile)
        if os.path.isfile(filename) and filename.startswith("./log"):
            os.remove(filename)


    # refresh_convertible_bond()
    
    refresh_optionBS_purebond_value_convertible_bond()
    
    # refresh_premium_rate_convertible_bond()
    # refresh_price_and_premium_rate_convertible_bond()
    # refresh_multifactor1_convertible_bond()
    # refresh_multifactor2_convertible_bond()

def main():
    
    main_function()
    schedule.every().day.at("07:00").do(main_function)  # 部署7：00执行更新数据任务
    while True:
        schedule.run_pending()


if __name__ == "__main__":
    main()
