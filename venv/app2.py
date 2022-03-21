
#pywin32 -> win32com


import win32com.client
import shutil
import time

# Based upon Code Sample from http://nbviewer.ipython.org/github/sanand0/ipython-notebooks/blob/master/Office.ipynb
#                   and from http://stackoverflow.com/questions/11832628/python-excel-macro-refresh

# Set Pathnames &amp; Filename (Use forward slashes / instead of backslashes \ in the paths)
SourcePathName = 'D:/OneDrive - Epiroc/CHLRUJ/Trabajo/Epiroc/PowerBI/Purchase Orders'
FileName = 'CHL Purchase Orders_Pending.xlsx'
print ('start get Excel')
# Open Excel
Application = win32com.client.Dispatch("Excel.Application")
print ('end get Excel')
# Show Excel. While this is not required, it can help with debugging
Application.Visible = 1
print ('start open File ')
# Open Your Workbook
Workbook = Application.Workbooks.open(SourcePathName + '/' + FileName)
print ('end open file')
print ('start refresh all')

#Application.OnTime = Now + TimeValue("00:05:00"), "SaveWb"

# Refesh All
Workbook.RefreshAll()
#Application.DoEvents()
time.sleep(30)

print ('end refresh all')


print ('start save file')

# Saves the Workbook
Workbook.Save()

print ('end save file')


print ('start quit file')

# Closes Excel
Application.Quit()
print ('end quit file')
