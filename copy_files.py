import os
import sys
import pathlib
import shutil

filter = ['rar']

source_dir = sys.argv[1]
target_dir = sys.argv[2]
print(f'Copying files from {source_dir} to {target_dir}')

for path in pathlib.Path(str(source_dir)).rglob('*.*'):
    if os.path.isdir(path):
        continue
    extension = pathlib.Path(path).suffix[1:].lower()
    if not extension in filter:
        continue
    
    source = str(path)
    print(source)
    target = source.replace(source_dir, target_dir)
    print(target)
        
    os.makedirs(os.path.dirname(target), exist_ok = True)
    shutil.copyfile(source, target)
        

input("Press Enter to continue...")
