#import modules
import pandas as pd
import pyodbc
import datetime

#these modules are not used, but are necessary to workaround PyInstaller error
import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy

#dictionary containing start and end date of each flyer month
flyer_dates = {"October 2019" : [datetime.date(2019,10,2), datetime.date(2019,10,29)], "November 2019" : [datetime.date(2019,10,30), datetime.date(2019,11,26)],
"December 2019" : [datetime.date(2019,11,27), datetime.date(2019,12,31)], "January 2020" : [datetime.date(2020,1,1), datetime.date(2020,1,28)], "February 2020" : [datetime.date(2020,1,29), datetime.date(2020,2,25)],
"March 2020" : [datetime.date(2020,2,26), datetime.date(2020,3,24)], "April 2020" : [datetime.date(2020,3,25), datetime.date(2020,4,28)], "May 2020" : [datetime.date(2020,4,29), datetime.date(2020,5,26)],
"June 2020" : [datetime.date(2020,5,27), datetime.date(2020,6, 30)], "July 2020" : [datetime.date(2020,7,1), datetime.date(2020,7,28)], "August 2020" : [datetime.date(2020,7,29), datetime.date(2020,8,25)],
"September 2020" : [datetime.date(2020,8,26), datetime.date(2020,9,29)], "October 2020" : [datetime.date(2020,9,30), datetime.date(2020,10,27)], "November 2020" : [datetime.date(2020,10,28), datetime.date(2020,11,24)],
"December 2020" : [datetime.date(2020,11,25), datetime.date(2020,12,29)], "January 2021" : [datetime.date(2021,12,30), datetime.date(2021,1,26)], "February 2021" : [datetime.date(2021,1,27), datetime.date(2021,2, 23)],
"March 2021" : [datetime.date(2021,2,24), datetime.date(2021,3, 30)], "April 2021" : [datetime.date(2021,3,31), datetime.date(2021,4,27)], "May 2021" : [datetime.date(2021,4,28), datetime.date(2021,5,25)], "June 2021" : [datetime.date(2021,5,26),
datetime.date(2021,6,29)], "July 2021" : [datetime.date(2021,6,30), datetime.date(2021,7,27)], "August 2021" : [datetime.date(2021,7,28), datetime.date(2021,8,31)], "September 2020" : [datetime.date(2021,9,1), datetime.date(2021,9,28)],
"October 2021" : [datetime.date(2021,9,29), datetime.date(2021,10,26)], "November 2021" : [datetime.date(2021,10,27), datetime.date(2021,11,30)], "December 2021" : [datetime.date(2021,12,1), datetime.date(2021,12,28)]}

#screen asking for user input. Debug mode allows for an existsing flyer date to be changed or a new on added
print("Flyer Performance Build")
print("1: Build Report")
print("2: Debug Mode")

entry = input("Select Option: ")

if entry == str(2):
    print("1: Change Flyer Date")
    print("2: Add New Flyer Date")
    debug_entry = input("Select Option: ")
    if debug_entry == str(1):
        debug_entry_1 = input("Which date would you like to change? (Example: October 2020) : ")
        print("Dates are case sensitive and in the format of 'YYYY,MM,DD'")
        print("Unlike previous formatting, no leading zeroes are accepted. 2019,01,21 would become: 2019,1,21")
        new_flyer_start = tuple(int(x.strip()) for x in input("New flyer start date = ").split(','))
        new_flyer_end = tuple(int(x.strip()) for x in input("New flyer end date = ").split(','))
        flyer_dates[debug_entry_1][0] = datetime.date(new_flyer_start[0],new_flyer_start[1],new_flyer_start[2])
        flyer_dates[debug_entry_1][1] = datetime.date(new_flyer_end[0],new_flyer_end[1],new_flyer_end[2])
        print("{0} dates have been changed". format(debug_entry_1))
    else:
        new_date = input("Enter the month and year (Example: July 2023) : ")
        print("Dates are case sensitive and in the format of 'YYYY,MM,DD'")
        print("Unlike previous formatting, no leading zeroes are accepted. 2019,01,21 would become: 2019,1,21")
        new_date_start = tuple(int(x.strip()) for x in input("Start date = ").split(','))
        new_date_end = tuple(int(x.strip()) for x in input("End date = ").split(','))
        new_date_start = datetime.date(new_date_start[0],new_date_start[1],new_date_start[2])
        new_date_end = datetime.date(new_date_end[0],new_date_end[1],new_date_end[2])
        flyer_dates[new_date] = [new_date_start, new_date_end]
        print("{0} has been added" .format(new_date))

#date selection
entry = input("Input a month & year (example: March 2021): ")

#progress bar 1
print(" Building Flyer Performance: [          ] 0/100", end='\r')

#declare today as the current date
today = datetime.date.today()

#determine flyer start and end, previous weeks start and end, as well as future weeks start and end.
flyer_start_date = flyer_dates[entry][0]
flyer_end_date = flyer_dates[entry][1]

difference = flyer_end_date - flyer_start_date + datetime.timedelta(days=1)

previous_weeks_start_date = flyer_start_date - difference
pwsd = str(previous_weeks_start_date) + ' 00:00:00.000'
previous_weeks_end_date = flyer_start_date - datetime.timedelta(days=1)
pwed = str(previous_weeks_end_date) + ' 23:59:59.000'

flyerwsd = str(flyer_start_date) + ' 00:00:00.000'
flyerwed = str(flyer_end_date) + ' 23:59:59.999'

future_weeks_start_date = flyer_end_date + datetime.timedelta(days=1)
fwsd = str(future_weeks_start_date) + ' 00:00:00.000'
future_weeks_end_date = flyer_end_date + difference
fwed = str(future_weeks_end_date) + ' 23:59:59.000'

#read buysheet information
buysheet = pd.read_excel('flyer.xlsx', dtype={0:'Int64'})

#sql connection parameters
conn = pyodbc.connect('Connection parameters here')
cursor = conn.cursor()

#sql query to pull item information for v_InventoryMaster
sql_query = """
SELECT INV_PK, INV_CPK, INV_ScanCode, t2.ASC_ScanCode
FROM v_InventoryMaster AS t1
LEFT JOIN AdditionalScanCodes AS t2 ON ASC_INV_FK = INV_PK AND ASC_INV_CFK = INV_CPK
WHERE t1.INV_STO_FK = 11
AND ISNUMERIC(INV_ScanCode) = 1
GROUP BY INV_PK, INV_CPK, INV_ScanCode, t2.ASC_ScanCode
ORDER BY INV_Scancode;
"""

query = pd.read_sql_query(sql_query, conn)

#progress bar 2
print (" Building Flyer Performance: [#         ] 10/100", end='\r')

item_info = pd.DataFrame(query)

#convert INV_ScanCode to numeric and delete duplicates
item_info['INV_ScanCode'] = item_info['INV_ScanCode'] = pd.to_numeric(item_info['INV_ScanCode'])
item_info = item_info.drop_duplicates(subset=['INV_ScanCode'], keep='first')

#set INV_ScanCode to index to map to buysheet
item_info.set_index('INV_ScanCode', inplace=True)
buysheet['INV_PK'] = buysheet.UPC.map(item_info.INV_PK)
buysheet['INV_CPK'] = buysheet.UPC.map(item_info.INV_CPK)

#set new dataframe equal to item_info sorted on numericn values of ASC_ScanCode
#use .copy() to ensure we don't get the dataframe slicing warning
item_info_alternate = item_info[pd.to_numeric(item_info['ASC_ScanCode'], errors='coerce').notnull()].copy()
#reset index
item_info_alternate.reset_index(inplace=True)
#convert ASC_ScanCode to numeric
item_info_alternate['ASC_ScanCode'] = pd.to_numeric(item_info_alternate['ASC_ScanCode'])
item_info_alternate['ASC_ScanCode'] = item_info_alternate['ASC_ScanCode'].astype('Int64')
#set ASC_ScanCode to index to map to buysheet
item_info_alternate.set_index('ASC_ScanCode', inplace=True)
buysheet['INV_PK2'] = buysheet.UPC.map(item_info_alternate.INV_PK)
buysheet['INV_CPK2'] = buysheet.UPC.map(item_info_alternate.INV_CPK)
#fill NAs with values from item_info_alternate
buysheet['INV_PK'] = buysheet['INV_PK'].fillna(buysheet['INV_PK2'])
buysheet['INV_CPK'] = buysheet['INV_CPK'].fillna(buysheet['INV_CPK2'])
#drop columns
buysheet = buysheet.drop(columns=['INV_PK2', 'INV_CPK2'])
#convert to numeric
buysheet['INV_PK'] = buysheet['INV_PK'].astype('int64')
buysheet['INV_CPK'] = buysheet['INV_CPK'].astype('int64')
#use INV_PK and INV_CPK to create a master key for each item. This will now be used instead of UPC, which is unreliable
buysheet['Combined'] = buysheet['INV_PK'].astype(str)+buysheet['INV_CPK'].astype(str)
buysheet['Combined'] = buysheet['Combined'].astype('int64')

#sql query using PromotionalSaleWorksheetData. This pulls each items cost and retail at the time of the worksheet being loaded.
sql_query = """
SELECT PSD_INV_FK, PSD_INV_CFK, PSD_CommitBasePrice1, PSD_CommitLastCost, t2.PSW_StartDate, t2.PSW_EndDate
FROM PromotionalSaleWorksheetData AS t1
LEFT JOIN PromotionalSaleWorksheet AS t2 ON PSD_PSW_FK = PSW_PK AND PSD_PSW_CFK = PSW_CPK
WHERE t2.PSW_ZON_FK = 6
AND t2.PSW_StartDate >= ?
AND t2.PSW_EndDate = ?
GROUP BY PSD_INV_FK, PSD_INV_CFK, PSD_CommitBasePrice1, PSD_CommitLastCost, t2.PSW_StartDate, t2.PSW_EndDate;
"""

query = pd.read_sql_query(sql_query, conn, params=(flyerwsd, flyerwed))

#progress bar 3
print (" Building Flyer Performance: [##        ] 20/100", end='\r')

pswd = pd.DataFrame(query)
#drop duplicates
pswd = pswd.drop_duplicates(subset=['PSD_INV_FK'], keep='first')
#create master key from PSD_INV_FK and PSD_INV_CFK
pswd['Combined'] = pswd['PSD_INV_FK'].astype(str)+pswd['PSD_INV_CFK'].astype(str)
pswd['Combined'] = pswd['Combined'].astype('int64')
#set as index and map to buysheet
pswd.set_index('Combined', inplace=True)
buysheet['Previous Cost'] = buysheet.Combined.map(pswd.PSD_CommitLastCost)
buysheet['Previous Retail'] = buysheet.Combined.map(pswd.PSD_CommitBasePrice1)

#the next three if statements will check if previous_weeks_end_date, flyer_end_date, and future_weeks_end_date
#are less than today's date. if they are, sales will be pulled. Else: fill column with 0
if previous_weeks_end_date < today:

    sql_query = """
    SELECT ITI_INV_FK, ITI_INV_CFK, SUM(TLI_Quantity) AS Quantity
    FROM v_TJTrans
    WHERE TRN_StartTime >= ?
    AND TRN_EndTime <= ?
    AND TLI_LIT_FK <> 4
    AND ITI_VOID = 0
    AND TRN_AllVoid = 0
    GROUP BY ITI_INV_FK, ITI_INV_CFK
    """

    query = pd.read_sql_query(sql_query, conn, params=(pwsd, pwed))
    #progress bar 4
    print(" Building Flyer Performance: [#####     ] 50/100", end='\r')
    previous_sales = pd.DataFrame(query)
    previous_sales['ITI_INV_FK'].fillna(99999, inplace=True)
    previous_sales['ITI_INV_CFK'].fillna(99999, inplace=True)
    previous_sales['ITI_INV_FK'] = previous_sales['ITI_INV_FK'].astype('int64')
    previous_sales['ITI_INV_CFK'] = previous_sales['ITI_INV_CFK'].astype('int64')
    previous_sales['Combined'] = previous_sales['ITI_INV_FK'].astype(str)+previous_sales['ITI_INV_CFK'].astype(str)
    previous_sales['Combined'] = previous_sales['Combined'].astype('int64')
    previous_sales.set_index('Combined', inplace=True)

    buysheet['Previous Sales'] = buysheet.Combined.map(previous_sales.Quantity)

else:

    buysheet['Previous Sales'] = 0

if flyer_end_date < today:

    query = pd.read_sql_query(sql_query, conn, params=(flyerwsd, flyerwed))
    #progress bar 5
    print(" Building Flyer Performance: [########  ] 80/100", end='\r')
    flyer_sales = pd.DataFrame(query)
    flyer_sales['ITI_INV_FK'].fillna(99999, inplace=True)
    flyer_sales['ITI_INV_CFK'].fillna(99999, inplace=True)
    flyer_sales['ITI_INV_FK'] = flyer_sales['ITI_INV_FK'].astype('int64')
    flyer_sales['ITI_INV_CFK'] = flyer_sales['ITI_INV_CFK'].astype('int64')
    flyer_sales['Combined'] = flyer_sales['ITI_INV_FK'].astype(str)+flyer_sales['ITI_INV_CFK'].astype(str)
    flyer_sales['Combined'] = flyer_sales['Combined'].astype('int64')
    flyer_sales.set_index('Combined', inplace=True)

    buysheet['Flyer Sales'] = buysheet.Combined.map(flyer_sales.Quantity)

else:

    buysheet['Flyer Sales'] = 0

if future_weeks_end_date < today:

    query = pd.read_sql_query(sql_query, conn, params=(fwsd, fwed))
    #progress bar 6
    print(" Building Flyer Performance: [######### ] 90/100", end='\r')
    future_weeks = pd.DataFrame(query)
    future_weeks['ITI_INV_FK'].fillna(99999, inplace=True)
    future_weeks['ITI_INV_CFK'].fillna(99999, inplace=True)
    future_weeks['ITI_INV_FK'] = future_weeks['ITI_INV_FK'].astype('int64')
    future_weeks['ITI_INV_CFK'] = future_weeks['ITI_INV_CFK'].astype('int64')
    future_weeks['Combined'] = future_weeks['ITI_INV_FK'].astype(str)+future_weeks['ITI_INV_CFK'].astype(str)
    future_weeks['Combined'] = future_weeks['Combined'].astype('int64')
    future_weeks.set_index('Combined', inplace=True)

    buysheet['Future Sales'] = buysheet.Combined.map(future_weeks.Quantity)

else:

    buysheet['Future Sales'] = 0

#fill NAs with 0
buysheet = buysheet.fillna(0)

#declare new columns for calculations
buysheet['Previous Margin'] = ((buysheet['Previous Retail']-buysheet['Previous Cost'])/buysheet['Previous Retail'])
buysheet['Previous Total'] = buysheet['Previous Retail'] * buysheet['Previous Sales']
buysheet['Flyer Margin'] = ((buysheet['Promo Retail']-buysheet['Promo Cost'])/buysheet['Promo Retail'])
buysheet['Flyer Total'] = buysheet['Promo Retail'] * buysheet['Flyer Sales']

buysheet['Flyer Movement Comparison'] = abs(buysheet['Flyer Sales'] - buysheet['Previous Sales'])
buysheet.loc[(buysheet['Flyer Sales'] == 0) & (buysheet['Previous Sales'] > 0), 'Flyer Movement Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Sales'] == 0) & (buysheet['Flyer Sales'] > 0), 'Flyer Movement Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Sales'] != 0) & (buysheet['Flyer Sales'] != 0), 'Flyer Movement Comparison Percentage'] = abs(((buysheet['Flyer Sales']-buysheet['Previous Sales'])/buysheet['Previous Sales']))

buysheet['Flyer Sales Comparison'] = abs(buysheet['Flyer Total']-buysheet['Previous Total'])
buysheet.loc[(buysheet['Flyer Total'] == 0) & (buysheet['Previous Total'] > 0), 'Flyer Sales Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Total'] == 0) & (buysheet['Flyer Total'] > 0), 'Flyer Sales Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Total'] != 0) & (buysheet['Flyer Total'] != 0), 'Flyer Sales Comparison Percentage'] = abs(((buysheet['Flyer Total']-buysheet['Previous Total'])/buysheet['Previous Total']))

buysheet['Future Total'] = buysheet['Future Sales'] * buysheet['Previous Retail']
buysheet['Future Movement Comparison'] = abs(buysheet['Future Sales'] - buysheet['Previous Sales'])
buysheet.loc[(buysheet['Future Sales'] == 0) & (buysheet['Previous Sales'] > 0), 'Future Movement Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Sales'] == 0) & (buysheet['Future Sales'] > 0), 'Future Movement Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Sales'] != 0) & (buysheet['Future Sales'] != 0), 'Future Movement Comparison Percentage'] = abs(((buysheet['Future Sales']-buysheet['Previous Sales'])/buysheet['Previous Sales']))

buysheet['Future Sales Comparison'] = abs(buysheet['Future Total']-buysheet['Previous Total'])
buysheet.loc[(buysheet['Future Total'] == 0 ) & (buysheet['Previous Total'] > 0), 'Future Sales Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Total'] == 0) & (buysheet['Future Total'] > 0), 'Future Sales Comparison Percentage'] = 1.0
buysheet.loc[(buysheet['Previous Total'] != 0) & (buysheet['Future Total'] != 0), 'Future Sales Comparison Percentage'] = abs(((buysheet['Future Total']-buysheet['Previous Total'])/buysheet['Previous Total']))

buysheet['Future Cost'] = buysheet['Previous Cost']
buysheet['Future Retail'] = buysheet['Previous Retail']
buysheet['Future Margin'] = buysheet['Previous Margin']

#drop unused columns
buysheet = buysheet.drop(columns=['INV_PK', 'INV_CPK', 'Combined'])

#rename headers
buysheet = buysheet[['UPC', 'Brand', 'Description', 'Size', 'Previous Cost', 'Previous Retail', 'Previous Margin', 'Previous Sales', 'Previous Total',
'Promo Cost', 'Promo Retail', 'Flyer Margin', 'Flyer Sales', 'Flyer Total', 'Flyer Movement Comparison', 'Flyer Movement Comparison Percentage',
'Flyer Sales Comparison', 'Flyer Sales Comparison Percentage', 'Future Cost', 'Future Retail', 'Future Margin', 'Future Sales', 'Future Total',
'Future Movement Comparison', 'Future Movement Comparison Percentage', 'Future Sales Comparison', 'Future Sales Comparison Percentage']]

#fill NAs once more
buysheet = buysheet.fillna(0)

#set Brand, Description, and Size columns to uppercase
buysheet['Brand'] = buysheet['Brand'].str.upper()
buysheet['Description'] = buysheet['Description'].str.upper()
buysheet['Size'] = buysheet['Size'].str.upper()

#write brand subtotals DataFrame
brandsub = pd.DataFrame(buysheet.groupby('Brand')['Previous Total'].sum())
brandsub['Flyer Total'] = buysheet.groupby('Brand')['Flyer Total'].sum()
brandsub['Flyer Comparison'] = abs(((brandsub['Flyer Total'] - brandsub['Previous Total'])/brandsub['Previous Total']))
brandsub['Future Total'] = buysheet.groupby('Brand')['Future Total'].sum()
brandsub['Future Comparison'] = abs(((brandsub['Future Total'] - brandsub['Previous Total'])/brandsub['Previous Total']))
brandsub = brandsub.reset_index()

#begin workbook
writer = pd.ExcelWriter("{0} Flyer Performance.xlsx" .format(entry), engine='xlsxwriter')
buysheet.to_excel(writer, sheet_name = 'Flyer Performance', startcol=0, startrow=1, index=False)
workbook = writer.book
worksheet = writer.sheets['Flyer Performance']
worksheet.set_zoom(70)

#declare formats
cell_format1 = workbook.add_format({'num_format' : '0 00000 00000 0', 'align' : 'center', 'border': 1})
cell_format2 = workbook.add_format({'border': 1})
description_header = workbook.add_format({'bold': True, 'align' : 'center'})
previous_bg_header = workbook.add_format({'bold': True, 'align' : 'center' , 'bg_color' : '#ddebf7', 'border' : 1})
flyer_bg_header = workbook.add_format({'bold': True, 'align' : 'center' , 'bg_color' : '#fff2cc' , 'border' : 1})
future_bg_header =  workbook.add_format({'bold': True, 'align' : 'center' , 'bg_color' : '#ebf1de' , 'border' : 1})
previous_bg_dollar = workbook.add_format({'num_format': '$0.00', 'align' : 'center', 'bg_color' : '#ddebf7', 'border': 1})
previous_bg_percentage = workbook.add_format({'num_format': '0.00%', 'align' : 'center', 'bg_color' : '#ddebf7' , 'border': 1})
align_center = workbook.add_format({'align' : 'center' , 'border': 1})
previous_bg = workbook.add_format({'align' : 'center', 'bg_color' : '#ddebf7' , 'border': 1})
flyer_bg = workbook.add_format({'align' : 'center', 'bg_color' : '#fff2cc' , 'border': 1})
flyer_bg_dollar = workbook.add_format({'num_format': '$0.00', 'align' : 'center', 'bg_color' : '#fff2cc' , 'border': 1})
flyer_bg_percentage = workbook.add_format({'num_format': '0.00%', 'align' : 'center', 'bg_color' : '#fff2cc' , 'border': 1})
future_bg = workbook.add_format({'align' : 'center', 'bg_color' : '#ebf1de' , 'border': 1})
future_bg_dollar = workbook.add_format({'num_format': '$0.00', 'align' : 'center', 'bg_color' : '#ebf1de' , 'border': 1})
future_bg_percentage = workbook.add_format({'num_format': '0.00%', 'align' : 'center', 'bg_color' : '#ebf1de' , 'border': 1})

#filter based off length of rows
df_length = str((len(buysheet)+2))
worksheet.autofilter('A2:AA'+df_length)

#set column widths and apply formatting, if necessary
worksheet.set_column('A:A', 19.43, cell_format1)
worksheet.set_column('B:B', 32.43, cell_format2)
worksheet.set_column('C:C', 77.14, cell_format2)
worksheet.set_column('D:D', 11.71, cell_format2)
worksheet.set_column('E:E', 17.43, previous_bg_dollar)
worksheet.set_column('F:F', 13.57, previous_bg_dollar)
worksheet.set_column('G:G', 15, previous_bg_percentage)
worksheet.set_column('H:H', 19.57, previous_bg)
worksheet.set_column('I:I', 18.86, previous_bg_dollar)
worksheet.set_column('J:J', 24.71, flyer_bg_dollar)
worksheet.set_column('K:K', 20.29, flyer_bg_dollar)
worksheet.set_column('L:L', 15, flyer_bg_percentage)
worksheet.set_column('M:M', 19.57, flyer_bg)
worksheet.set_column('N:N', 18.86, flyer_bg_dollar)
worksheet.set_column('O:O', 19.57, flyer_bg)
worksheet.set_column('P:P', 22.29, flyer_bg_percentage)
worksheet.set_column('Q:Q', 18.86, flyer_bg_dollar)
worksheet.set_column('R:R', 21.29, flyer_bg_percentage)
worksheet.set_column('S:S', 17.43, future_bg_dollar)
worksheet.set_column('T:T', 13.57, future_bg_dollar)
worksheet.set_column('U:U', 15, future_bg_percentage)
worksheet.set_column('V:V', 19.57, future_bg)
worksheet.set_column('W:W', 18.86, future_bg_dollar)
worksheet.set_column('X:X', 19.57, future_bg)
worksheet.set_column('Y:Y', 22.29, future_bg_percentage)
worksheet.set_column('Z:Z',18.86, future_bg_dollar)
worksheet.set_column('AA:AA', 21.29, future_bg_percentage)

#write headers and merge
merge_format_previous = workbook.add_format({ 'bold': 1, 'border': 1, 'align': 'center', 'fg_color': '#9bc2e6'})
merge_format_flyer = workbook.add_format({ 'bold': 1, 'border': 1, 'align': 'center', 'fg_color': '#ffd966'})
merge_format_future = workbook.add_format({ 'bold': 1, 'border': 1, 'align': 'center', 'fg_color': '#c4d79b'})
previous_bg_header.set_font_size(14)
flyer_bg_header.set_font_size(14)
future_bg_header.set_font_size(14)
merge_format_previous.set_font_size(16)
merge_format_flyer.set_font_size(16)
merge_format_future.set_font_size(16)
worksheet.merge_range('E1:I1', 'Previous Weeks', merge_format_previous)
worksheet.merge_range('J1:N1', 'Flyer Weeks', merge_format_flyer)
worksheet.merge_range('O1:R1', 'Comparisons', merge_format_flyer)
worksheet.merge_range('S1:W1', 'Future Weeks', merge_format_future)
worksheet.merge_range('X1:AA1', 'Comparisons', merge_format_future)

#green and red formats for comparison highlighting
red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

worksheet.conditional_format('O3:O'+df_length, {'type': 'formula', 'criteria': 'M3-H3 < 0', 'format': red_format})
worksheet.conditional_format('O3:O'+df_length, {'type': 'formula', 'criteria': 'M3-H3 > 0', 'format': green_format})
worksheet.conditional_format('P3:P'+df_length, {'type': 'formula', 'criteria': 'M3-H3 < 0', 'format': red_format})
worksheet.conditional_format('P3:P'+df_length, {'type': 'formula', 'criteria': 'M3-H3 > 0', 'format': green_format})
worksheet.conditional_format('Q3:Q'+df_length, {'type': 'formula', 'criteria': 'N3-I3 < 0', 'format': red_format})
worksheet.conditional_format('Q3:Q'+df_length, {'type': 'formula', 'criteria': 'N3-I3 > 0', 'format': green_format})
worksheet.conditional_format('R3:R'+df_length, {'type': 'formula', 'criteria': 'N3-I3 < 0', 'format': red_format})
worksheet.conditional_format('R3:R'+df_length, {'type': 'formula', 'criteria': 'N3-I3 > 0', 'format': green_format})
worksheet.conditional_format('X3:X'+df_length, {'type': 'formula', 'criteria': 'V3-H3 < 0', 'format': red_format})
worksheet.conditional_format('X3:X'+df_length, {'type': 'formula', 'criteria': 'V3-H3 > 0', 'format': green_format})
worksheet.conditional_format('Y3:Y'+df_length, {'type': 'formula', 'criteria': 'V3-H3 < 0', 'format': red_format})
worksheet.conditional_format('Y3:Y'+df_length, {'type': 'formula', 'criteria': 'V3-H3 > 0', 'format': green_format})
worksheet.conditional_format('Z3:Z'+df_length, {'type': 'formula', 'criteria': 'W3-I3 < 0', 'format': red_format})
worksheet.conditional_format('Z3:Z'+df_length, {'type': 'formula', 'criteria': 'W3-I3 > 0', 'format': green_format})
worksheet.conditional_format('AA3:AA'+df_length, {'type': 'formula', 'criteria': 'W3-I3 < 0', 'format': red_format})
worksheet.conditional_format('AA3:AA'+df_length, {'type': 'formula', 'criteria': 'W3-I3 > 0', 'format': green_format})

#rename headers. This is necessary so the formatting takes effect
worksheet.write(1,0, 'UPC', description_header)
worksheet.write(1,1, 'Brand', description_header)
worksheet.write(1,2, 'Description', description_header)
worksheet.write(1,3, 'Size', description_header)
worksheet.write(1,4, 'Cost', previous_bg_header)
worksheet.write(1,5, 'Retail', previous_bg_header)
worksheet.write(1,6, 'Margin', previous_bg_header)
worksheet.write(1,7, 'Movement', previous_bg_header)
worksheet.write(1,8, 'Total Sales', previous_bg_header)
worksheet.write(1,9, 'Cost', flyer_bg_header)
worksheet.write(1,10, 'Retail', flyer_bg_header)
worksheet.write(1,11, 'Margin', flyer_bg_header)
worksheet.write(1,12, 'Movement', flyer_bg_header)
worksheet.write(1,13, 'Total Sales', flyer_bg_header)
worksheet.write(1,14, 'Movement', flyer_bg_header)
worksheet.write(1,15, 'Movement %', flyer_bg_header)
worksheet.write(1,16, 'Sales Total', flyer_bg_header)
worksheet.write(1,17, 'Sales Total %', flyer_bg_header)
worksheet.write(1,18, 'Unit Cost', future_bg_header)
worksheet.write(1,19, 'Retail', future_bg_header)
worksheet.write(1,20, 'Margin', future_bg_header)
worksheet.write(1,21, 'Movement', future_bg_header)
worksheet.write(1,22, 'Total Sales', future_bg_header)
worksheet.write(1,23, 'Movement', future_bg_header)
worksheet.write(1,24, 'Movement %', future_bg_header)
worksheet.write(1,25, 'Sales Total', future_bg_header)
worksheet.write(1,26, 'Sales Total %', future_bg_header)

#hide rows that do not contain values
worksheet.set_default_row(hide_unused_rows=True)
worksheet.freeze_panes(0, 4)

#begin sheet 2
brandsub.to_excel(writer, sheet_name = 'Brand Subtotals', startcol=0, startrow=1, index=False)
worksheet2 = writer.sheets['Brand Subtotals']
worksheet2.set_zoom(85)

#filter based off length of rows
df2_length = str((len(brandsub)+2))
worksheet2.autofilter('A2:F2'+df2_length)

#declare formats
merge_format_previous_brd = workbook.add_format({ 'bold': 1, 'border': 1, 'align': 'center', 'fg_color': '#9bc2e6'})
merge_format_flyer_brd = workbook.add_format({ 'bold': 1, 'border': 1, 'align': 'center', 'fg_color': '#ffd966'})
merge_format_future_brd = workbook.add_format({ 'bold': 1, 'border': 1, 'align': 'center', 'fg_color': '#c4d79b'})
description_header_brd = workbook.add_format({'bold': True, 'align' : 'center'})
previous_bg_header_brd = workbook.add_format({'bold': True, 'align' : 'center' , 'bg_color' : '#ddebf7', 'border' : 1})
flyer_bg_header_brd = workbook.add_format({'bold': True, 'align' : 'center' , 'bg_color' : '#fff2cc' , 'border' : 1})
future_bg_header_brd =  workbook.add_format({'bold': True, 'align' : 'center' , 'bg_color' : '#ebf1de' , 'border' : 1})
cell_format1_brd = workbook.add_format({'align' : 'left', 'border': 1})
previous_bg_brd = workbook.add_format({'align' : 'center', 'bg_color' : '#ddebf7' , 'border': 1, 'num_format': 44})
flyer_bg_brd = workbook.add_format({'align' : 'center', 'bg_color' : '#fff2cc' , 'border': 1, 'num_format': 44})
future_bg_brd = workbook.add_format({'align' : 'center', 'bg_color' : '#ebf1de' , 'border': 1, 'num_format': 44})

#write headers and merge
worksheet2.write(0,1, 'Previous Weeks', merge_format_previous_brd)
worksheet2.merge_range('C1:D1', 'Flyer Weeks', merge_format_flyer_brd)
worksheet2.merge_range('E1:F1', 'Future Weeks', merge_format_future_brd)
worksheet2.write(1,0, 'Brand',description_header_brd)
worksheet2.write(1,1, 'Brand Subtotals', previous_bg_header_brd)
worksheet2.write(1,2, 'Brand Subtotals', flyer_bg_header_brd)
worksheet2.write(1,3, 'Percentage', flyer_bg_header_brd)
worksheet2.write(1,4, 'Brand Subtotals', future_bg_header_brd)
worksheet2.write(1,5, 'Percentage', future_bg_header_brd)

#set column widths and formats
worksheet2.set_column('A:A', 29.14, cell_format1_brd)
worksheet2.set_column('B:B', 19.14, previous_bg_brd)
worksheet2.set_column('C:C', 19.14, flyer_bg_brd)
worksheet2.set_column('D:D', 15, flyer_bg_percentage)
worksheet2.set_column('E:E', 19.14, future_bg_brd)
worksheet2.set_column('F:F', 15, future_bg_percentage)

merge_format_previous_brd.set_font_size(12)
merge_format_flyer_brd.set_font_size(12)
merge_format_future_brd.set_font_size(12)

#set conditional formats
worksheet2.conditional_format('D3:D'+df2_length, {'type': 'formula', 'criteria': 'C3-B3 < 0', 'format': red_format})
worksheet2.conditional_format('D3:D'+df2_length, {'type': 'formula', 'criteria': 'C3>B3 > 0', 'format': green_format})
worksheet2.conditional_format('F3:F'+df2_length, {'type': 'formula', 'criteria': 'E3-B3 < 0', 'format': red_format})
worksheet2.conditional_format('F3:F'+df2_length, {'type': 'formula', 'criteria': 'E3>B3 > 0', 'format': green_format})

#hide unused rows
worksheet2.set_default_row(hide_unused_rows=True)
#save
writer.save()

#progress bar 7
print(" Building Flyer Performance: [##########] 100/100", end='\r')
print("")
print("Complete!")
