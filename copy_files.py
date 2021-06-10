import os
import sys
import pathlib
import shutil

filter = ['msg']

source_dir = 'X:\\ZZ\\IF\\' + sys.argv[1]
target_dir = 'X:\\' + sys.argv[1]
print(f'Copying files from {source_dir} to {target_dir}')

count = 0
for path in pathlib.Path(str(source_dir)).rglob('*.*'):
    if os.path.isdir(path):
        continue
    extension = pathlib.Path(path).suffix[1:].lower()
    if not extension in filter:
        continue
        
    source = str(path)
    target = source.replace(source_dir, target_dir)
    print(target)
    
    if os.path.exists(str(target)[:-3] + 'pdf'):
        continue
        
    os.makedirs(os.path.dirname(target), exist_ok = True)
    shutil.copyfile(source, target)
    os.remove(source)
    
    if os.path.exists(target + '.txt'):
        os.remove(target + '.txt')
        
    count += 1
        
print(f'copied {count} files')
input("Press Enter to continue...")
