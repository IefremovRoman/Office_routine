# Imports
import os
import pprint
import openpyxl
import sys
from itertools import islice
import pandas as pd

# Path description and preparing part
path = input('Type your folder:')
cur_dir = os.listdir(path)
pp = pprint.PrettyPrinter()
pp.pprint(cur_dir)

file = 'bunch.xlsx'
xls = openpyxl.load_workbook(path+'\\'+file)
sheets = xls.get_sheet_names()
pp.pprint(sheets)

# Preparing for Pandas.DataFrame conversion and printing the data as DataFrame
def excel_to_df(excel_file,active_sheet):
  sheet = excel_file[active_sheet]
  data = sheet.values

  # Indicate the columns in the sheet values
  cols = next(data)[1:]

  # Convert your data to a list
  data = list(data)

  # Read in the data at index 0 for the indices
  idx = [r[0] for r in data]

  # Slice the data at index 1 
  data = (islice(r, 1, None) for r in data)

  # Make your DataFrame
  df = pd.DataFrame(data, index=idx, columns=cols)		
  return df

main_df = excel_to_df(xls,xls.active.title)[0:0]
for sheet in sheets:
  df = excel_to_df(xls,sheet)
  main_df = main_df.append(df)

main_df.drop(main_df.columns[[-1,-2]], axis=1, inplace=True)
main_df = main_df.drop_duplicates(subset=main_df.columns[0], inplace=False)
main_df = main_df.sort_values(main_df.columns[3])
main_df.to_excel(path+'\\'+'bunch1.xlsx',sheet_name='Total')

print('\n')
sys.exit('Total sheet was made succesfully!')