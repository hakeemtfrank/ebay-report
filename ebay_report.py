# Ebay Report Script #
# v.1 12/1/17 #

import os
import time
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings("ignore")


# Import Dataset from comp_data/ 
print("\n AVAILABLE DATABASES: \n ")
print(str(os.listdir("comp_data/")))

my_ebay_user = input("\n What is your ebay username? \n \n ")
ebay_user = input("\n What is your competition\'s ebay username? (check spelling!) \n \n ")  # NOTEBOOK IMPORT: ebay user name
my_datapath = "comp_data/" + my_ebay_user + '.csv'
comp_datapath = "comp_data/" + ebay_user + '.csv'  # IMPORTANT! : This is the format of goofbids ebay_user.csv files

print("\n Spying...")
my_df = pd.read_csv(my_datapath, encoding="latin1", parse_dates=['End Date'])
comp_df = pd.read_csv(comp_datapath, encoding="latin1", parse_dates=['End Date'])

##################
# CLEAN DATASETS #
##################

dropcol = ['Buyer', 'Bid Count']

comp_df = comp_df.drop(dropcol, axis=1)
comp_df["Price"] = comp_df["Price"].replace('[\$,]', '', regex=True)
comp_df["Shipping"] = comp_df["Shipping"].replace('[\$,]', '', regex=True)
comp_df["Price"] = pd.to_numeric(comp_df["Price"].get_values())
comp_df["Shipping"] = pd.to_numeric(comp_df["Shipping"].get_values())
comp_df["Total Sales"] = comp_df["Price"] * comp_df["No. Sold"]

my_df = my_df.drop(dropcol, axis=1)
my_df["Price"] = my_df["Price"].replace('[\$,]', '', regex=True)
my_df["Shipping"] = my_df["Shipping"].replace('[\$,]', '', regex=True)
my_df["Price"] = pd.to_numeric(my_df["Price"].get_values())
my_df["Shipping"] = pd.to_numeric(my_df["Shipping"].get_values())
my_df["Total Sales"] = my_df["Price"] * my_df["No. Sold"]

rel_sales = comp_df[["Title", "No. Sold", "Price", "Total Sales"]]  # Relevant Columns

print("...")

#######################
# PERFORMANCE METRICS #
#######################

my_totalsales = my_df["Total Sales"].sum()
my_salesvol = my_df["No. Sold"].sum()
my_itemct = len(my_df)
my_unsold = (my_df["No. Sold"] == 0).sum()
my_invratio = (my_salesvol / my_df["Quantity"].sum())

comp_totalsales = comp_df["Total Sales"].sum()
comp_salesvol = comp_df["No. Sold"].sum()
comp_itemct = len(comp_df)
comp_unsold = (comp_df["No. Sold"] == 0).sum()
comp_invratio = (comp_salesvol / comp_df["Quantity"].sum())

# TABLE FOR COMP TOP TEN ITEMS BY QT #
qtsales = (rel_sales.sort_values(["No. Sold"], ascending=False).reset_index(drop=True))[:10]
qtsales.index = np.arange(1, len(qtsales) + 1)


# TABLE FOR COMP TOP TEN ITEMS BY TOT SALES #
totsales = (rel_sales.sort_values(["Total Sales"], ascending=False).reset_index(drop=True))[:10]
totsales.index = np.arange(1, len(totsales) + 1)

print("...")

#########################
# CONVERT DF TO REPORTS #
#########################

print("\n Reporting...")

file_path = ebay_user + "_report" + "/"  # new folder per ebay user
directory = os.path.dirname(file_path)  

if not os.path.exists(directory): 
  os.makedirs(directory)

xlsfile = file_path + ebay_user + '_analysis.xlsx'  
writer = pd.ExcelWriter(xlsfile, engine='xlsxwriter')  

# Fill in a sheet with tsales and qsales dataframes
workbook = writer.book
overviewsheet = workbook.add_worksheet('Overview')

totsales.to_excel(writer, sheet_name='Top 10 by Total Sales', startrow=6)  # WARNING: Intl Table Placement
qtsales.to_excel(writer, sheet_name='Top 10 by Total Sales', startrow=19)  # WARNING: Intl Table Placement

rankingsheet = writer.sheets['Top Items']  

print("...")

####################
# FORMATTING RULES #
####################

rankingsheet.set_zoom(120)
overviewsheet.set_zoom(120)

headerformat = workbook.add_format()
headerformat.set_font_size(13)
headerformat.set_align('center')
headerformat.set_align('vcenter')

subheaderformat = workbook.add_format()
subheaderformat.set_font_size(12)
subheaderformat.set_align('center')
subheaderformat.set_align('vcenter')

titlesform = workbook.add_format()
titlesform.set_font_size(12)
titlesform.set_align('center')
titlesform.set_align('vcenter')
titlesform.set_bold()

toplink_format = workbook.add_format({
    'font_color': 'blue',
    'underline': 1,
    'font_size': 13,
    'align': 'center'
})
toplink_format.set_align('vcenter')

money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
pct_fmt = workbook.add_format({'num_format': '0.00%'})
bold = workbook.add_format({'bold': True})

color_fmt_red = workbook.add_format({'font_color': '#9C0006'})  # dark red
color_fmt_grn = workbook.add_format({'font_color': '#006100'})  # dark green

print("...")

#############################
# OVERVIEW SHEET FORMATTING #
#############################

overviewsheet.conditional_format('C6:D6', {'type': 'top',
                                           'value': '1',
                                           'format': color_fmt_grn})

overviewsheet.conditional_format('C6:D6', {'type': 'bottom',
                                           'value': '1',
                                           'format': color_fmt_red})

overviewsheet.conditional_format('C7:D7', {'type': 'top',
                                           'value': '1',
                                           'format': color_fmt_grn})

overviewsheet.conditional_format('C7:D7', {'type': 'bottom',
                                           'value': '1',
                                           'format': color_fmt_red})

overviewsheet.conditional_format('C8:D8', {'type': 'top',
                                           'value': '1',
                                           'format': color_fmt_red})

overviewsheet.conditional_format('C8:D8', {'type': 'bottom',
                                           'value': '1',
                                           'format': color_fmt_grn})

overviewsheet.conditional_format('C9:D9', {'type': 'top',
                                           'value': '1',
                                           'format': color_fmt_grn})

overviewsheet.conditional_format('C9:D9', {'type': 'bottom',
                                           'value': '1',
                                           'format': color_fmt_red})

for i in range(0, 15):
  overviewsheet.set_row(i, 20)

overviewsheet.set_column('A:A', 3)  # Index
overviewsheet.set_column('B:B', 21)  # Descriptors
overviewsheet.set_column('C:D', 30)  # No. Sold

overviewsheet.write(3, 2, my_ebay_user, subheaderformat)
overviewsheet.write(3, 3, ebay_user, subheaderformat)

overviewsheet.write(5, 1, "Total Reported Sales")
overviewsheet.write(5, 2, my_totalsales, money_fmt)
overviewsheet.write(5, 3, comp_totalsales, money_fmt)

overviewsheet.write(6, 1, "Total Reported Sales Volume")
overviewsheet.write(6, 2, my_salesvol)
overviewsheet.write(6, 3, comp_salesvol)

overviewsheet.write(4, 1, "Items for Sale")
overviewsheet.write(4, 2, my_itemct)
overviewsheet.write(4, 3, comp_itemct)

overviewsheet.write(7, 1, "Unsold Items")
overviewsheet.write(7, 2, my_unsold)
overviewsheet.write(7, 3, comp_unsold)

overviewsheet.write(8, 1, "% Sold from Inventory")
overviewsheet.write(8, 2, my_invratio, pct_fmt)
overviewsheet.write(8, 3, comp_invratio, pct_fmt)

##############################
# TOP ITEMS SHEET FORMATTING #
##############################

rankingsheet.conditional_format('E8:E17', {'type': 'top',
                                           'value': '6',
                                           'format': color_fmt_grn})

rankingsheet.conditional_format('E8:E17', {'type': 'bottom',
                                           'value': '4',
                                           'format': color_fmt_red})

rankingsheet.conditional_format('C21:C30', {'type': 'top',
                                            'value': '6',
                                            'format': color_fmt_grn})

rankingsheet.conditional_format('C21:C30', {'type': 'bottom',
                                            'value': '4',
                                            'format': color_fmt_red})


# Top Items Sheet - Header Row Widths
header_rows = [1, 2, 3, 5, 18]
for i in header_rows:
  rankingsheet.set_row(i, 22)

# Top Items Sheet - Special Column Widths
rankingsheet.set_column('A:A', 3)  # Index
rankingsheet.set_column('B:B', 65)  # Item Name
rankingsheet.set_column('C:C', 8)  # No. Sold
rankingsheet.set_column('D:D', 11, money_fmt)  # Price
rankingsheet.set_column('E:E', 14, money_fmt)  # Total Sales

# Headers for Top Items Report
rankheader = 'Performance Overview'
analysisdate = 'Updated: ' + time.strftime('%A, %x')
store_link = 'https://www.ebay.com/usr/' + ebay_user
rankingsheet.write(1, 1, rankheader, headerformat)  # WRITE Title
rankingsheet.write_url(2, 1, store_link, toplink_format, string=ebay_user)  # Store URL
rankingsheet.write(3, 1, analysisdate, subheaderformat)  # WRITE Date
rankingsheet.write(5, 1, 'Top 10 Items by Total Sales', titlesform)
rankingsheet.write(18, 1, 'Top 10 Items by Quantity Sold', titlesform)


writer.save()

print("\n Finished! \n")
print("Your report is in: ../ebay_analysis/" + file_path + " \n ")
