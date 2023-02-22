import pandas as pd
import openpyxl
import os


# Load the data into a pandas dataframe
df = pd.read_csv('inventory.csv')

# Find items that are below 5 in quantity and need to be ordered again
need_order = df[df['quantity'] < 5]

# Find items that have more than 10 in quantity and another group that has over 30 in quantity
group_10 = df[(df['quantity'] > 10) & (df['quantity'] <= 30)]
group_30 = df[df['quantity'] > 30]

# Find the top 10 worst items sold
worst_items = df.nsmallest(10, 'quantity')

# Create a new Excel workbook and write the results to a new sheet
with pd.ExcelWriter('inventory_report_reorder.xlsx') as writer:
    need_order.to_excel(writer, sheet_name='Need Order')
    group_10.to_excel(writer, sheet_name='Group 10+')
    group_30.to_excel(writer, sheet_name='Group 30+')
    worst_items.to_excel(writer, sheet_name='Worst Items Sold')

os.system('start excel.exe inventory_report_reorder.xlsx')
