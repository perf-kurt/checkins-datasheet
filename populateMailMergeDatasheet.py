# Iterate files in /exports
import os
from openpyxl import load_workbook

# Load mail merge datasheet template
mail_merge_data_wb = load_workbook(filename='utils/Mail Merge Data Template.xlsx')
data_entry_wb = load_workbook(filename='imports/Data Entry.xlsx')
mm_sheet = mail_merge_data_wb.active

# Read data from each sheet in data entry file
for sheet in data_entry_wb.worksheets:
    # Read data from each row in the sheet
    for row in sheet.iter_rows(min_row=7, values_only=True):
        if row[0] is not None:  # Make sure data exists
            mm_sheet['A' + str(mm_sheet.max_row + 1)] = sheet.title.split(' - ')[0]  # File name without extension
            mm_sheet['B' + str(mm_sheet.max_row)] = row[0]
            mm_sheet['C' + str(mm_sheet.max_row)] = row[1]
            mm_sheet['D' + str(mm_sheet.max_row)] = row[2]
            mm_sheet['E' + str(mm_sheet.max_row)] = row[3]
            mm_sheet['F' + str(mm_sheet.max_row)] = row[4]
            mm_sheet['G' + str(mm_sheet.max_row)] = row[5]
            mm_sheet['H' + str(mm_sheet.max_row)] = row[6]
            mm_sheet['I' + str(mm_sheet.max_row)] = row[7]
            mm_sheet['J' + str(mm_sheet.max_row)] = sheet.title.split(' - ')[1]  # Tier
            mm_sheet['K' + str(mm_sheet.max_row)] = sheet['C6'].value  # Attr 1
            mm_sheet['L' + str(mm_sheet.max_row)] = sheet['D6'].value  # Attr 2
            mm_sheet['M' + str(mm_sheet.max_row)] = sheet['E6'].value  # Attr 3
            mm_sheet['N' + str(mm_sheet.max_row)] = sheet['F6'].value  # Attr 4
            mm_sheet['O' + str(mm_sheet.max_row)] = sheet['G6'].value  # Attr 5
        
    print('Processed sheet: ' + sheet.title)

mail_merge_data_wb.save('Mail Merge Datasheet (output).xlsx')