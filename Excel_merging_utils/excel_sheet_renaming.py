# Imports
import os
import pprint
import openpyxl

# Path specifying
path = input('Type your dir: ')
current_dir = os.listdir(path)

# Logging
pp = pprint.PrettyPrinter()
print('\n')
print(path)
print('\n')
pp.pprint(current_dir)

# Working on sheets
for file in current_dir:
  print('\n'*2)
  print(file)
  # Work with file if it's not as below
  if file != 'bunch.xlsx':
    filename = path+'\\'+file
    print(filename)
    print(f'Opening {file}....')
    xls = openpyxl.load_workbook(filename)					
    print(type(xls))
    cur_xls_sheets = xls.active

    print(f'In {file} we have some sheets:')
    pp.pprint(cur_xls_sheets)

    print(f'Current sheet is {cur_xls_sheets}')
    
    # Renaming sheets
    cur_xls_sheets.title = file[:file.rfind('.')]
    
    # Saving
    xls.save(filename)
    print(f'{file} saved succesfully!')