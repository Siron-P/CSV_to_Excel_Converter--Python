import openpyxl
import pandas as pd

df = pd.read_csv('ABCD.CSV')

df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
dfm= df.drop_duplicates()
df = df.fillna('')
for col in df.select_dtypes(include = 'str').columns:
    df[col] = df[col].str.strip()

df.to_excel('ABCD.xlsx',index=False)

wb = openpyxl.load_workbook('ABCD.xlsx')
ws = wb.active

for cell in ws[1]:
    cell.font = openpyxl.styles.Font(bold= True)

wb.save('ABCD.xlsx')
print("Done!")