import os
import sys
import pathlib
import binascii

file_extensions = []

bad_files = ['doc', 'docm', 'dot', 'dotm', 'odt', 'xls', 'xlsm', 'xlsb', 'xlt', 'xltm', 'ods', 'ppt', 'pptm', 'pot', 'potm', 'pps', 'ppsm', 'odp']

source_dir = sys.argv[1]
print(f'Processing folder: {source_dir}')

for path in pathlib.Path(str(source_dir)).rglob('*.*'):
    if os.path.isdir(path):
        continue
    extension = pathlib.Path(path).suffix[1:].lower()
    if not extension in file_extensions:
        file_extensions.append(extension)
        
    #with open (path, "rb") as myfile:
    #    header = myfile.read(4)
    #    header = str(binascii.hexlify(header))
     #   #print(f'{header} -> {extension}')
        
file_extensions.sort() 
print(file_extensions)

for bad_file in bad_files:
    if bad_file in file_extensions:
        print(f'ERROR: found {bad_file} extension!')

input("Press Enter to continue...")
