import openpyxl

file = input('Type your file with path: ')

print(openpyxl.load_workbook(file).sheetnames)