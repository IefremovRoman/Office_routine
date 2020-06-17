# Imports
import os
import pyexcel as p
import pprint

# Parameters
path = input('Type your dir: ')
remove_option = input('Delete xls files? (y/n) ')

# Your currnet directory with files
current_dir = os.listdir(path)
pp = pprint.PrettyPrinter()
pp.pprint(current_dir)

# Looking inside current directory
for file in current_dir:							
  print('\n'*2)
  print(file)
  # specified files, which do not required to be transformed
  if file != '____.xlsx':
    filename = path+'\\'+file
    print(filename)
    # Transforming operation by renaming
    p.save_book_as(file_name=filename,
                   dest_file_name=filename[:filename.rfind('.')]+'.xlsx')
    print('Saved as %s' % filename[:filename.rfind('.')]+'.xlsx')
    if remove_option == 'y':
      os.remove(filename)
      print(f'{file} was removed!')