import pandas as pd
import glob
import openpyxl

files = glob.glob("invoices/*.xlsx")
print(files)

for file in files:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    print(df)