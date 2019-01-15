import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = openpyxl.load_workbook("D:\\OneDrive - Epiroc\\CHLRUJ\\Trabajo\\Epiroc\\PowerBI\\Purchase Orders\\CHL Purchase Orders_Pending.xlsx")


ws = wb.worksheets[0]


tab = wb._pivots[0]




ws._
#tab = Table(displayName="TableData")




