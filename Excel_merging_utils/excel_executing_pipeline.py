import os
import subprocess
from openpyxl import Workbook

# Save your current work dir
cur_dir = os.getcwd()

# Call scripts
subprocess.call("python xls_to_xlsx.py", shell=True)
subprocess.call("python excel_sheet_renaming.py", shell=True)

# Change dir
os.chdir(input('Type your dir: '))
work_dir = os.getcwd()

# Create empty file (required for next steps)
wb = Workbook()
wb.save("bunch.xlsx")

# Back to working dir, where to call other scripts
os.chdir(cur_dir)
subprocess.call("python excel_mergin.py", shell=True)
subprocess.call("python excel_merging_to_1_sheet_unique.py", shell=True)

#Renaming
os.chdir(work_dir)
new_name = os.path.basename(os.getcwd())
new_name.replace(' ', '_')
os.rename('bunch.xlsx', new_name+'.xlsx')