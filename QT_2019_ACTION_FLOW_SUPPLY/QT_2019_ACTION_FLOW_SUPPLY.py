import xlwings as xw
import pandas as pd
import win32api



excel_file_directory = "I:\\github_repos\\work-6\\QT_2019_ACTION_FLOW_SUPPLY\\QT_2019_ACTION_FLOW_SUPPLY.xlsm"

# Requied columns
columns = ['COMPANY NAME', 'LOCATION', 'PHONE', 'ADDRESS', 'Second ship address', 'CONTACT 1', 
            'CONTACT 2', 'CONTACT 3', 'CONTACT 4', 'CONTACT 5', 'CONTACT 6']


# ===============================================================INITIALIZATION==================================================================================
"""
    Initialize workbook and sheet variables
"""
wb = xw.Book.caller()
# wb.sheets['test'].range("A1").value = "Hello xlwings!"

sht_quotation = wb.sheets['Quotation']
sht_sales_order = wb.sheets['Sales Order']
sht_customer = wb.sheets['Customers']
sht_test = wb.sheets['test']    # test sheet

# =================================================================RUN Button====================================================================================
def quotation_run():
    # Parse the 'Customers' data into dataframe
    df = pd.ExcelFile(excel_file_directory).parse('Customers')
    df = df[columns]
    # sht_test.clear()        # clear content and formatting
    # sht_test.range('A1').options(index= False).value = df
    # sht_test.range('A1:Z1000000').autofit()

    #----------------------------------------------------------------------------------------------------------------------------------------------------
    search1_company_in = sht_quotation.range('H5').value     # input -- to be entered into search box in 'Quotation' sheet

    df_search1 = df.loc[df['COMPANY NAME'].isin([search1_company_in])]      # search for company input

    if df_search1.empty == False:       # check if the dataframe is not empty
        if len(df_search1['COMPANY NAME'].tolist()) > 1:    # check if the dataframe by company_name has more than 1 row

            # populate the Location column cells with location data
            sht_quotation.range('Z2:AZ2').clear_contents      # clear content only 
            sht_quotation.range('Z2').value = df_search1['LOCATION'].tolist()

            if sht_quotation.range('I5').value is None:     # check if the cell 'I5' is empty
                win32api.MessageBox(wb.app.hwnd, "Since, more than 1 element is found, so please enter the 'Location' as 2nd parameter", "Search by Company")

            else:       # if the Location box is filled with location data
                search1_location_in = sht_quotation.range('I5').value
                df_search1_location = df_search1.loc[df_search1['LOCATION'].isin([search1_location_in])]    # search for location input

                if df_search1_location.empty == False:      # check if location based dataframe is not empty
                    # display data
                    sht_quotation.range('B10').value = df_search1_location['CONTACT 1'].tolist()[0]       # contact
                    sht_quotation.range('B11').value = df_search1_location['COMPANY NAME'].tolist()[0]       # company
                    sht_quotation.range('B12').value = df_search1_location['LOCATION'].tolist()[0]       # location
                    sht_quotation.range('B13').value = df_search1_location['ADDRESS'].tolist()[0]       # address
                    sht_quotation.range('B14').value = df_search1_location['PHONE'].tolist()[0]       # phone
                else:
                    # ignoring the case where 'I5' is empty. Basically, here it is not available in the Location list items, so don't prompt any dialog
                    if sht_quotation.range('I5').value is None:      
                        pass
                    else:
                        win32api.MessageBox(wb.app.hwnd, "SORRY! The Location name doesn't exist.", "Search by Company")

        else:
            # display data
            sht_quotation.range('B10').value = df_search1['CONTACT 1'].tolist()[0]       # contact
            sht_quotation.range('B11').value = df_search1['COMPANY NAME'].tolist()[0]       # company
            sht_quotation.range('B12').value = df_search1['LOCATION'].tolist()[0]       # location
            sht_quotation.range('B13').value = df_search1['ADDRESS'].tolist()[0]       # address
            sht_quotation.range('B14').value = df_search1['PHONE'].tolist()[0]       # phone

    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Company name doesn't exist.", "Search by Company")

# ==========================================================RESET Button=======================================================================================
def quotation_reset():
    sht_quotation.range('Z2:AZ2').clear_contents()
    sht_quotation.range('H5').clear_contents()
    sht_quotation.range('I5').clear_contents()
