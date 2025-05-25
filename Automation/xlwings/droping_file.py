import pandas as pd
import xlwings as xl
import os

file_loc = r'C:\Users\kiran\OneDrive\Desktop\Excel-Automation-With-Python\Automation\Excel_Files\first.xlsx'

if not os.path.exists(file_loc):
    print("Bro file is not exists")
else:
    print('Bro file found let me delete it')
    os.remove(file_loc)
    print("bro done")
print('see you :)')
