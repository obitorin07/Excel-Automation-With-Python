import pandas as pd
import xlwings as xl

file = r"G:\Projects\Pizza sales\pizza_sales.xlsm"
df = pd.read_excel(file)
print(df.shape)

print(xl.Book(file))
print(xl.Range("A2:C50").value)
