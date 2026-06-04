import os
import openpyxl
import pandas as pd
import argparse
import logging

#For logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("log.txt"),
        logging.StreamHandler()
    ]
)

#for CLI argument -argparse
parser = argparse.ArgumentParser()

parser.add_argument('--input',required = True, help='Path to CSV  File')
parser.add_argument('--output',required = True, help='Path to save Excel File')
args = parser.parse_args()

#for file not found error
if not os.path.exists(args.input):
    print(f"Error: File {args.input} not found")
    logging.error(f"Error: File {args.input} not found")
    exit()

if not args.input.lower().endswith('.csv'):
    logging.error(f"Error: File {args.input} is not in csv format.")
    exit()

#for cleaning and normalizing data -pandas
df = pd.read_csv(args.input)

df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
df = df.drop_duplicates()
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

#for formatting the excel file -openpyxl
wb = openpyxl.load_workbook(args.output)
ws = wb.active

for col in ws.columns:
    max_length = max(len(str(cell.value)) for cell in col if cell.value)
    ws.column_dimensions[col[0].column_letter].width = max_length + 4


for cell in ws[1]:
    cell.font = openpyxl.styles.Font(bold= True)

wb.save('ABCD.xlsx')
logging.info(f"{args.input} file converted to {args.output} file.")