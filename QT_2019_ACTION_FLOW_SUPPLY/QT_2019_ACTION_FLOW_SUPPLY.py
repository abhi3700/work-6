import xlwings as xw
import pandas as pd
import win32api



excel_file_directory = "I:\\github_repos\\work-6\\QT_2019_ACTION_FLOW_SUPPLY\\QT_2019_ACTION_FLOW_SUPPLY.xlsm"

# Requied columns
columns = ['COMPANY NAME', 'LOCATION', 'PHONE', 'ADDRESS', 'Second ship address', 'CONTACT 1', 'CONTACT 2', 'CONTACT 3', 'CONTACT 4', 'CONTACT 5', 'CONTACT 6']

def main():
    wb = xw.Book.caller()
    # wb.sheets['test'].range("A1").value = "Hello xlwings!"

    sht_quotation = wb.sheets['Quotation']
    sht_sales_order = wb.sheets['Sales Order']
    sht_customer = wb.sheets['Customers']
    sht_test = wb.sheets['test']    # test sheet


    df = pd.ExcelFile(excel_file_directory).parse('Customers')
    df = df[columns]
    # sht_test.clear()        # clear content and formatting
    # sht_test.range('A1').options(index= False).value = df
    # sht_test.range('A1:Z1000000').autofit()

    #----------------------------------------------------------------------------------------------------------------------------------------------------
    search1_in = sht_quotation.range('H5').value     # input -- to be entered into search box in 'Quotation' sheet

    df_search1 = df.loc[df['COMPANY NAME'].isin([search1_in])]

    if df['COMPANY NAME'].isin([search1_in]) == True:
        df_search1 = df.loc[df['COMPANY NAME'].isin([search1_in])]

        search1_contact = df_search1['CONTACT 1'].tolist()[0]       # contact
        search1_company = df_search1['COMPANY NAME'].tolist()[0]       # company
        search1_location = df_search1['LOCATION'].tolist()[0]       # location
        search1_address = df_search1['ADDRESS'].tolist()[0]       # address
        search1_phone = df_search1['PHONE'].tolist()[0]       # address
        
        #----------------------------------------------------------------------------------------------------------------------------------------------------
        # display data
        sht_quotation.range('B10').value = search1_contact
        sht_quotation.range('B11').value = search1_company
        sht_quotation.range('B12').value = search1_location
        sht_quotation.range('B13').value = search1_address
        sht_quotation.range('B14').value = search1_phone

    else:
        win32api.MessageBox(wb.app.hwnd, "Search by Company", "Check Sub-category", )



@xw.func
def hello(name):
    return "hello {0}".format(name)
