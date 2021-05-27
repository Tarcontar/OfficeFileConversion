import os
import sys
import pathlib

print(f'deleting all files of type _tmp.zip')

dir = sys.argv[1]

for path in pathlib.Path(str(dir)).rglob('*.*'):
    extension = pathlib.Path(path).suffix[1:].lower()
    if not extension in ['zip']:
        continue
    
    name = os.path.basename(path)[:-4]
    if not name == '_tmp':
        continue
        
    print (path)
    os.remove(path)
        
input("Press Enter to continue...")
