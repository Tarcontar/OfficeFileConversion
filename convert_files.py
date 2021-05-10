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


#source_dir = str(current_dir) + '/source'
source_dir = 'C:\\C'
issue_target_dir = 'C:\\CI'
legacy_target_dir = 'C:\\CB'

logfile = open('C:\\log.txt', 'a')

def get_magic(path):
    with open (path, 'rb') as myfile:
        header = myfile.read(4)
        return str(binascii.hexlify(header))


def copy_file(source, target):
    os.makedirs(target.replace(os.path.basename(target), ''), exist_ok = True)
    shutil.copyfile(source, target)


def handle_error(path):
    error_msg = f'ERROR: could not convert \'{path}\' \n'
    print(error_msg)
    logfile.write(error_msg)
    placeholder = open(path + '.txt', 'w')
    placeholder.write('file could not be converted')
    placeholder.close()
    relpath = path.replace(str(source_dir), '')
    newPath = issue_target_dir + relpath
    copy_file(path, newPath)
    os.remove(path)
    
    
count = 0

for path in pathlib.Path(source_dir).rglob('*.*'):
    extension = pathlib.Path(path).suffix[1:].lower()

    path = str(path)
    print(os.path.basename(path))
    if extension in ['docx', 'doc', 'docm', 'dot', 'dotm', 'odt']:
        ff = DOCX_FILE_FORMAT
        if path.endswith('docx'):
            if '504b0304' in get_magic(path):
                continue
            else:
                print('WARNING: fake file detected')
                os.rename(path, path[:-1])
                path = path[:-1]
        
        print(path)
        
        try:
            doc = word.Documents.Open(path, ConfirmConversions=False, Visible=False)
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

        word.ActiveDocument.SaveAs(new_path, ff)
        doc.Close(False)
        copy_file(path, path.replace(source_dir, legacy_target_dir))
        os.remove(path)
        count += 1
        
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
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            wb = excel.Workbooks.Open(path)
            wb.Application.DisplayAlerts = False
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
        
        wb.SaveAs(new_path, FileFormat=ff, ConflictResolution=2)
        wb.Close()
        copy_file(path, path.replace(source_dir, legacy_target_dir))
        os.remove(path)
        count += 1
        
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
            presentation = ppt.Presentations.Open(path, WithWindow=False)
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
        copy_file(path, path.replace(source_dir, legacy_target_dir))
        os.remove(path)
        count += 1
       
try:     
    word.Application.Quit()
except:
    pass
    
try:     
    excel.Application.Quit()
except:
    pass
    
try:     
    ppt.Quit()
except:
    pass

logfile.close()
print(f'converted {count} files')
input('Press Enter to continue...')

