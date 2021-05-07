import pathlib
import binascii

file_extensions = []

current_dir = pathlib.Path(__file__).parent.absolute()

for path in pathlib.Path(str(current_dir) + '/source').rglob('*.*'):
    extension = pathlib.Path(path).suffix
    if not extension in file_extensions:
        file_extensions.append(extension)
        
    with open (path, "rb") as myfile:
        header = myfile.read(4)
        header = str(binascii.hexlify(header))
        print(f'{header} -> {extension}')
        
file_extensions.sort() 
print(file_extensions)

input("Press Enter to continue...")
