import xlwings as xl
import pandas as pd
import os

file = r'C:\Users\kiran\OneDrive\Desktop\Excel-Automation-With-Python\Automation\Excel_Files\first1.xlsx'

if not os.path.exists(file):
    print('File not exist bro we can create')
    x = xl.Book()
    x.save(file)
    print('done bro check again' + " " +file)
    x.close()
else:
    print('File Exist Bro')


