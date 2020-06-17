# Importing libraries

import os
import pprint
import openpyxl

# Path descriptor part
path = input('Type your folder:')
cur_dir = os.listdir(path)
pp = pprint.PrettyPrinter()
pp.pprint(cur_dir)

file = input('Type your excel file (with extension):')
xls = openpyxl.load_workbook(path+'\\'+file)

sheets = xls.get_sheet_names()
pp.pprint(sheets)

# Working with cells in all sheets
for sheet in sheets:
  cur_sheet = xls[sheet]
  column_from = input('Specify column to delete from: ')
  column_to = input('Specify column to delete to: ')
  cur_sheet.delete_cols(column_from,column_to)
  print(type(sheet))

# Saving as new file
xls.save(path+'\\'+'new_excel_file.xlsx')