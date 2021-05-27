import os
import sys
import pathlib
import binascii
import subprocess

file_extensions = []

word_filter = ['doc', 'docm', 'dot', 'dotm', 'odt']
excel_filter = ['xls', 'xlsm', 'xlsb', 'xlt', 'xltm', 'ods', 'xlam']
ppt_filter = ['ppt', 'pptm', 'pot', 'potm', 'pps', 'ppsm', 'odp']
outlook_filter = ['msg']
malicious_filter = ['osd', 'py', 'exe', 'msi', 'bat', 'reg', 'pol', 'ps1', 'psm1', 'psd1', 'ps1xml', 'pssc', 'psrc', 'cdxml']

source_dir = sys.argv[1]
print(f'Processing folder: {source_dir}')

bad_files = word_filter + excel_filter + ppt_filter + outlook_filter + malicious_filter

for path in pathlib.Path(str(source_dir)).rglob('*.*'):
    if os.path.isdir(path):
        continue
    extension = pathlib.Path(path).suffix[1:].lower()
    if not extension in file_extensions:
        file_extensions.append(extension)
        
    if extension in bad_files:
        print(path)
        subprocess.Popen(f'explorer /select,{str(path)}')
        input()
        
file_extensions.sort() 
print(file_extensions)

for bad_file in bad_files:
    if bad_file in file_extensions:
        print(f'ERROR: found {bad_file} extension!')

input("Press Enter to continue...")
