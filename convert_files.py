import os
import re
import pathlib
import binascii
import shutil
import win32com.client as win32
from win32com.client import constants

import win32com
print(win32com.__gen_path__)

DOCX_FILE_FORMAT = 12
DOTX_FILE_FORMAT = 14
PPTX_FILE_FORMAT = 24
POTX_FILE_FORMAT = 26
PPSX_FILE_FORMAT = 28
XLSX_FILE_FORMAT = 51
XLTX_FILE_FORMAT = 54


current_dir = pathlib.Path(__file__).parent.absolute()
print(f'processing all [\'doc\', \'docm\', \'odt\', \'xls\', \'xlsm\', \'xlsb\', \'ods\', \'ppt\', \'pptm\', \'odp\'] files in \'{current_dir}\'')
print('do NOT close any opening office application windows (minimize them instead)')

word = win32.gencache.EnsureDispatch('Word.Application')
word.DisplayAlerts = False
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.DisplayAlerts = False
ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
ppt.DisplayAlerts = constants.ppAlertsNone


def get_magic(path):
    with open (path, 'rb') as myfile:
        header = myfile.read(4)
        return str(binascii.hexlify(header))

def handle_error(path):
    #logfile = open(pathlib.Path(str(current_dir) + 'log.txt'), 'w')
    print(f'ERROR: could not convert \'{path}\'')
    placeholder = open(path + '.txt', 'w')
    placeholder.write('file could not be converted')
    placeholder.close()
    relpath = path.replace(str(current_dir), '')
    newPath = 'C:\\FCI' + relpath
    os.makedirs(newPath.replace(os.path.basename(newPath), ''), exist_ok = True)
    shutil.copyfile(path, newPath)
    os.remove(path)
    

for path in pathlib.Path(str(current_dir) + '/source').rglob('*.*'):
    extension = pathlib.Path(path).suffix[1:].lower()

    path = str(path)
    print(os.path.basename(path))
    if extension in ['docx', 'doc', 'docm', 'dot', 'dotm', 'odt']:
        ff = DOCX_FILE_FORMAT
        if path.endswith('docx'):
            if '504b0304' in get_magic(path):
                continue
            else:
                print('fake file detected')
                os.rename(path, path[:-1])
                path = path[:-1]
        
        print(path)
        
        try:
            doc = word.Documents.Open(path)
        except:
            handle_error(path)
            continue
        doc.Activate()
        
        if path.endswith('dot') or path.endswith('dotm'):
            ff = DOTX_FILE_FORMAT

        if path.endswith('odt'):
            new_path = path[:-3] + 'docx'
        elif path.endswith('docm'):
            new_path = path[:-1] + 'x'
        elif path.endswith('dotm'):
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'

        word.ActiveDocument.SaveAs(new_path, FileFormat=ff)
        doc.Close(False)
        os.remove(path)
        
    elif extension in ['xlsx', 'xls', 'xlsm', 'xlsb', 'xlt', 'xltm', 'ods']:
        ff = XLSX_FILE_FORMAT
        if path.endswith('xlsx'):
            if '504b0304' in get_magic(path):
                continue
            else:
                print('fake file detected')
                os.rename(path, path[:-1])
                path = path[:-1]
        
        print(path)
        
        try:
            wb = excel.Workbooks.Open(path)
        except:
            handle_error(path)
            continue
            
        if path.endswith('xlt') or path.endswith('xltm'):
            ff = XLTX_FILE_FORMAT
            
        if path.endswith('ods'):
            new_path = path[:-3] + 'xlsx'
        elif path.endswith('xlsm'):
            new_path = path[:-1] + 'x'
        elif path.endswith('xlsb'):
            new_path = path[:-1] + 'x'
        elif path.endswith('xltm'):
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'
        
        wb.SaveAs(new_path, FileFormat = ff)
        wb.Close()
        os.remove(path)
        
    elif extension in ['pptx', 'ppt', 'pptm', 'pot', 'potm', 'pps', 'ppsm', 'odp']:
        ff = PPTX_FILE_FORMAT
        if path.endswith('pptx'):
            if '504b0304' in get_magic(path):
                continue
            else:
                print('fake file detected')
                os.rename(path, path[:-1])
                path = path[:-1]
        
        print(path)
        
        try:
            presentation = ppt.Presentations.Open(path, WithWindow = False)
        except:
            handle_error(path)
            continue
            
        if path.endswith('pot') or path.endswith('potm'):
            ff = POTX_FILE_FORMAT
        elif path.endswith('pps') or path.endswith('ppsm'):
            ff = PPSX_FILE_FORMAT
            
        
        if path.endswith('odp'):
            new_path = path[:-3] + 'pptx'
        elif path.endswith('pptm'):
            new_path = path[:-1] + 'x'
        elif path.endswith('potm'):
            new_path = path[:-1] + 'x'
        elif path.endswith('ppsm'):
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'
            
        presentation.SaveAs(new_path, ff)
        presentation.Close()
        os.remove(path)
       
try:     
    word.Application.Quit()
    excel.Application.Quit()
    ppt.Quit()
except:
    pass

input("Press Enter to continue...")

