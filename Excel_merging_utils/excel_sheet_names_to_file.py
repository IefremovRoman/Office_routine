import openpyxl

file = input('Type your file with path: ')

lst = openpyxl.load_workbook(file).sheetnames
with open('files.txt','w+') as file:
    for i in sorted(lst):
        file.write(i+ '\n')