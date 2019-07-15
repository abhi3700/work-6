import xlwings as xw    # for integrating python with Excel
import pandas as pd     # for DataFrame
import numpy as np      # for calling NaN values 
import win32api         # for message box



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

#----------------------------------------------------------------------------------------------------------------------------------------------------
# Parse the 'Customers' data into dataframe
df = pd.ExcelFile(excel_file_directory).parse('Customers')
df = df[columns]
# sht_test.clear()        # clear content and formatting
# sht_test.range('A1').options(index= False).value = df
# sht_test.range('A1:Z1000000').autofit()

def clear_output_cells():
    # Clear contents in the output cells
    sht_quotation.range('B10:B14').clear_contents()
    sht_quotation.range('D10:D14').clear_contents()
    sht_quotation.range('D13').clear_contents()

# =================================================================Search-1: RUN Button====================================================================================
def quotation_search1_run():

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
                    # Clear contents in the output cells
                    clear_output_cells()
                    # display data
                    # 'BILL TO' column 
                    contact1_loc = df_search1_location['CONTACT 1'].tolist()[0]       # contact 1
                    contact2_loc = df_search1_location['CONTACT 2'].tolist()[0]       # contact 2
                    contact3_loc = df_search1_location['CONTACT 3'].tolist()[0]       # contact 3
                    contact4_loc = df_search1_location['CONTACT 4'].tolist()[0]       # contact 4
                    contact5_loc = df_search1_location['CONTACT 5'].tolist()[0]       # contact 5
                    contact6_loc = df_search1_location['CONTACT 6'].tolist()[0]       # contact 6
                    if contact1_loc is np.nan:
                        if contact2_loc is np.nan:
                            if contact3_loc is np.nan:
                                if contact4_loc is np.nan:
                                    if contact5_loc is np.nan:
                                        if contact6_loc is np.nan:
                                            sht_quotation.range('B10').value = ""
                                            sht_quotation.range('D10').value = ""
                                        else:
                                            sht_quotation.range('B10').value = contact6_loc
                                            sht_quotation.range('D10').value = contact6_loc
                                    else:
                                        sht_quotation.range('B10').value = contact5_loc
                                        sht_quotation.range('D10').value = contact5_loc
                                else:
                                    sht_quotation.range('B10').value = contact4_loc
                                    sht_quotation.range('D10').value = contact4_loc
                            else:
                                sht_quotation.range('B10').value = contact3_loc
                                sht_quotation.range('D10').value = contact3_loc
                        else:
                            sht_quotation.range('B10').value = contact2_loc
                            sht_quotation.range('D10').value = contact2_loc
                    else:
                        sht_quotation.range('B10').value = contact1_loc
                        sht_quotation.range('D10').value = contact1_loc
                    sht_quotation.range('B11').value = df_search1_location['COMPANY NAME'].tolist()[0]       # company
                    sht_quotation.range('B12').value = df_search1_location['LOCATION'].tolist()[0]       # location
                    sht_quotation.range('B13').value = df_search1_location['ADDRESS'].tolist()[0]       # address
                    sht_quotation.range('B14').value = df_search1_location['PHONE'].tolist()[0]       # phone

                    # 'SHIP TO' column 
                    sht_quotation.range('D11').value = df_search1_location['COMPANY NAME'].tolist()[0]       # company
                    sht_quotation.range('D12').value = df_search1_location['LOCATION'].tolist()[0]       # location
                    ship_address_location = df_search1_location['Second ship address'].tolist()[0]      # shipping address
                    if ship_address_location is np.nan:
                        sht_quotation.range('D13').value = df_search1_location['ADDRESS'].tolist()[0]
                    else:
                        sht_quotation.range('D13').value = ship_address_location
                    sht_quotation.range('D14').value = df_search1_location['PHONE'].tolist()[0]       # phone
                else:
                    # ignoring the case where 'I5' is empty. Basically, here it is not available in the Location list items, so don't prompt any dialog
                    if sht_quotation.range('I5').value is None:      
                        pass
                    else:
                        win32api.MessageBox(wb.app.hwnd, "SORRY! The Location name doesn't exist.", "Search by Company")

        else:
            # Clear contents in the output cells
            clear_output_cells()

            # display data
            # 'BILL TO' column 
            contact1 = df_search1['CONTACT 1'].tolist()[0]       # contact 1
            contact2 = df_search1['CONTACT 2'].tolist()[0]       # contact 2
            contact3 = df_search1['CONTACT 3'].tolist()[0]       # contact 3
            contact4 = df_search1['CONTACT 4'].tolist()[0]       # contact 4
            contact5 = df_search1['CONTACT 5'].tolist()[0]       # contact 5
            contact6 = df_search1['CONTACT 6'].tolist()[0]       # contact 6
            if contact1 is np.nan:
                if contact2 is np.nan:
                    if contact3 is np.nan:
                        if contact4 is np.nan:
                            if contact5 is np.nan:
                                if contact6 is np.nan:
                                    sht_quotation.range('B10').value = ""
                                    sht_quotation.range('D10').value = ""
                                else:
                                    sht_quotation.range('B10').value = contact6
                                    sht_quotation.range('D10').value = contact6
                            else:
                                sht_quotation.range('B10').value = contact5
                                sht_quotation.range('D10').value = contact5
                        else:
                            sht_quotation.range('B10').value = contact4
                            sht_quotation.range('D10').value = contact4
                    else:
                        sht_quotation.range('B10').value = contact3
                        sht_quotation.range('D10').value = contact3
                else:
                    sht_quotation.range('B10').value = contact2
                    sht_quotation.range('D10').value = contact2
            else:
                sht_quotation.range('B10').value = contact1
                sht_quotation.range('D10').value = contact1
            sht_quotation.range('B11').value = df_search1['COMPANY NAME'].tolist()[0]       # company
            sht_quotation.range('B12').value = df_search1['LOCATION'].tolist()[0]       # location
            sht_quotation.range('B13').value = df_search1['ADDRESS'].tolist()[0]       # address
            sht_quotation.range('B14').value = df_search1['PHONE'].tolist()[0]       # phone
            # 'SHIP TO' column 
            sht_quotation.range('D11').value = df_search1['COMPANY NAME'].tolist()[0]       # company
            sht_quotation.range('D12').value = df_search1['LOCATION'].tolist()[0]       # location
            ship_address = df_search1['Second ship address'].tolist()[0]        # shipping address
            if ship_address is np.nan:
                sht_quotation.range('D13').value = df_search1['ADDRESS'].tolist()[0]
            else:
                sht_quotation.range('D13').value = ship_address
            sht_quotation.range('D14').value = df_search1['PHONE'].tolist()[0]       # phone

    elif sht_quotation.range('H5').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter Company Name in the search box", "Search by Company")

    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Company name doesn't exist.", "Search by Company")

# ==========================================================Search-1: RESET Button=======================================================================================
def quotation_search1_reset():
    sht_quotation.range('Z2:AZ2').clear_contents()
    sht_quotation.range('H5').clear_contents()
    sht_quotation.range('I5').clear_contents()

# ==========================================================Search-2: RUN Button=======================================================================================
def quotation_search2_run():
    search2_contact_in = sht_quotation.range('H20').value     # input -- to be entered into search box in 'Quotation' sheet

    df_search2 = df.loc[df['CONTACT 1'].isin([search2_contact_in])]      # search for contact input
    if df_search2.empty == False:
        # Clear contents in the output cells
        clear_output_cells()

        # Display data
        sht_quotation.range('B10').value = df_search2['CONTACT 1'].tolist()[0]       # contact
        sht_quotation.range('B11').value = df_search2['COMPANY NAME'].tolist()[0]       # company
        sht_quotation.range('B12').value = df_search2['LOCATION'].tolist()[0]       # location
        sht_quotation.range('B13').value = df_search2['ADDRESS'].tolist()[0]       # address
        sht_quotation.range('B14').value = df_search2['PHONE'].tolist()[0]       # phone
        # 'SHIP TO' column 
        sht_quotation.range('D10').value = df_search2['CONTACT 1'].tolist()[0]       # contact
        sht_quotation.range('D11').value = df_search2['COMPANY NAME'].tolist()[0]       # company
        sht_quotation.range('D12').value = df_search2['LOCATION'].tolist()[0]       # location
        ship_address_search2 = df_search2['Second ship address'].tolist()[0]        # shipping address
        if ship_address_search2 is np.nan:
            sht_quotation.range('D13').value = df_search2['ADDRESS'].tolist()[0]
        else:
            sht_quotation.range('D13').value = ship_address_search2
        sht_quotation.range('D14').value = df_search2['PHONE'].tolist()[0]       # phone
    elif sht_quotation.range('H20').value is None:
        win32api.MessageBox(wb.app.hwnd, "Please, enter Contact Name in the search box", "Search by Contact")

    else:
        win32api.MessageBox(wb.app.hwnd, "SORRY! The Contact name doesn't exist.", "Search by Contact")


# ==========================================================Search-2: RESET Button=======================================================================================
def quotation_search2_reset():
    sht_quotation.range('H20').clear_contents()
