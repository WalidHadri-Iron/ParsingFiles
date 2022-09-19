import xlwings as xw
import pandas as pd
import time
import os
import datetime

path_excel_file = r'C:/Users/walid/Desktop/XLWINGS/MASTER_WORKBOOK.xlsm' ### Change path to master workbook
sheets_to_extract = {"MODEL":["LOAD","DATA","METRICS","MODEL","GRAPHS","EBITDA"], 
                     "IncomeStatement":["DATA"],
                     "BalanceSheet":["DATA"],
                     "CashFlow":["DATA"]}
naming_extract = {"MODEL":-1,"IncomeStatement":"INCOME","BalanceSheet":"BALANCE","CashFlow":"CASHFLOW"}


#Getting the tickers
app = xw.App(visible=False)
app.display_alerts = False
wb = xw.Book(path_excel_file, update_links=False, ignore_read_only_recommended=True)
tickers_available = xw.books['MASTER_WORKBOOK.xlsm'].sheets['Admin_Control'].range('L:L').value[1:]
tickers_new = []
for ticker in tickers_available:
    if ticker:
        tickers_new.append(ticker)
    else:
        break
tickers_available = tickers_new
wb.close()
app.quit()
name_folder = datetime.datetime.now().strftime("%B")[:4].upper() + "{:02d}".format(datetime.datetime.now().day) + str(datetime.datetime.now().year)[2:]
try:
    os.mkdir(os.getcwd()+'\\EXPORT_'+name_folder)
except:
    pass
#loop through tickers
for ticker in tickers_available:
    list_loop = ["MODEL", "IncomeStatement", "BalanceSheet", "CashFlow"]
    for element in list_loop:
        
        sheets = sheets_to_extract[element]
        naming = naming_extract[element]

    
        if element == 'MODEL':
            app = xw.App()
            app.display_alerts = False
            wb = xw.Book(path_excel_file, update_links=False, ignore_read_only_recommended=True) 
            time.sleep(6)
            xw.books['MASTER_WORKBOOK.xlsm'].sheets['Admin_Control'].range('TICKER').value = ticker
            xw.books['MASTER_WORKBOOK.xlsm'].sheets['Admin_Control'].range('CTRL_STMT').value = element
            wb.api.Application.Run("Retrieve")
            wb.sheets["Admin_Control"].delete()
            wb.save('out.xlsx')
            wb.close()
            app.quit()

        else:
            app = xw.App()
            app.display_alerts = False
            wb = xw.Book(path_excel_file, update_links=False, ignore_read_only_recommended=True) 
            time.sleep(6)
            xw.books['MASTER_WORKBOOK.xlsm'].sheets['Admin_Control'].range('TICKER').value = ticker
            xw.books['MASTER_WORKBOOK.xlsm'].sheets['Admin_Control'].range('CTRL_STMT').value = element
            wb.api.Application.Run("Retrieve")
            app.display_alerts = False
            new_wb = xw.Book("out.xlsx", update_links=False, ignore_read_only_recommended=True) 
            sheet_name = sheets[0]
            sht = wb.sheets[sheet_name]
            sht.copy(after=new_wb.sheets[-1], name=naming)
            new_wb.save('out.xlsx')
            new_wb.close()
            wb.close()
            time.sleep(3)
            app.quit()
    
    app = xw.App()
    app.display_alerts = False
    wb = xw.Book('out.xlsx', update_links=False) 
    wb.save(os.getcwd()+'\\EXPORT_'+name_folder+'\\{}.xls'.format(ticker))
    wb.close()
    app.quit()