import os
import sys
import pathlib
import binascii

file_extensions = []


source_dir = sys.argv[1]
print(f'Processing folder: {source_dir}')

for path in pathlib.Path(str(source_dir)).rglob('*.*'):
    if os.path.isdir(path):
        continue
    extension = pathlib.Path(path).suffix
    if not extension in file_extensions:
        file_extensions.append(extension)
        
    #with open (path, "rb") as myfile:
    #    header = myfile.read(4)
    #    header = str(binascii.hexlify(header))
     #   #print(f'{header} -> {extension}')
        
file_extensions.sort() 
print(file_extensions)

input("Press Enter to continue...")
