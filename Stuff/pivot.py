from openpyxl import load_workbook

wb = load_workbook(filename = 'SR.xlsx')
sheet_ranges = wb['SR - By Analyst - Completed']
print(sheet_ranges['B10'].value)