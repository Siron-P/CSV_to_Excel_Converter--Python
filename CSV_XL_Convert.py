import os
import openpyxl
import pandas as pd
import argparse

parser = argparse.ArgumentParser()

parser.add_argument('--input',required = True, help='Path to CSV  File')
parser.add_argument('--output',required = True, help='Path to save Excel File')
args = parser.parse_args()

if not os.path.exists(args.input):
    print(f"Error: File {args.input} not found")
    exit()

if not args.input.lower().endswith('.csv'):
    print(f"Error: File {args.input} is not in csv format.")

df = pd.read_csv(args.input)

df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
dfm= df.drop_duplicates()
df = df.fillna('')
for col in df.select_dtypes(include = 'str').columns:
    df[col] = df[col].str.strip()

for col in df.columns:
    if 'date' in col or 'time' in col:
        try:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        except Exception:
            pass

df.to_excel(args.output,index=False)

wb = openpyxl.load_workbook('ABCD.xlsx')
ws = wb.active

for col in ws.columns:
    max_length = max(len(str(cell.value)) for cell in col if cell.value)
    ws.column_dimensions[col[0].column_letter].width = max_length + 4


for cell in ws[1]:
    cell.font = openpyxl.styles.Font(bold= True)

wb.save('ABCD.xlsx')
print("Done!")