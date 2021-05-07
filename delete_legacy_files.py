import os
import pathlib

files_to_delete = ['doc', 'odt', 'docm', 'xls', 'xlsm', 'xlsb', 'ods', 'ppt', 'pptm', 'odp']

print(f'deleting all files of type {files_to_delete}')

current_dir = pathlib.Path(__file__).parent.absolute()

for path in pathlib.Path(str(current_dir)).rglob('*.*'):
    extension = path.name.split('.')[-1]
    if extension in files_to_delete: # todo: remove or rename these
        os.remove(path)
        
input("Press Enter to continue...")
