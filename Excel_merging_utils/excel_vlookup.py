
import os
import openpyxl
from pprint import pprint

pprint(os.listdir(input('Type your source file path: ')))
souce_file = input('Choose your source file: ')
print('File {} has be chosen as source'.format(souce_file))

workfolder = os.chdir(input('Type your files` path: '))
print('='*32)
print('Your work folder:')
pprint(os.listdir(workfolder))

