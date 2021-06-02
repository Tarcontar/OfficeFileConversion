import os
import pathlib

file_filter = ['docx', 'xlsx', 'pptx', 'dotx', 'xltx', 'potx', 'ppsx', 'txt']

print(f'deleting all files of type {file_filter}')

current_dir = pathlib.Path(__file__).parent.absolute()

for path in pathlib.Path(str(current_dir)).rglob('*.*'):
    extension = pathlib.Path(path).suffix[1:].lower()
    if extension in file_filter:
        os.remove(path)
        
input("Press Enter to continue...")
