import os
import re
import pathlib
import shutil
import win32com.client as win32
from win32com.client import constants

XLSX_FILE_FORMAT = 51
PPTX_FILE_FORMAT = 24
PPSX_FILE_FORMAT = 24

current_dir = pathlib.Path(__file__).parent.absolute()
print(f'processing all [\'doc\', \'docm\', \'odt\', \'xls\', \'xlsm\', \'xlsb\', \'ods\', \'ppt\', \'pptm\', \'odp\'] files in \'{current_dir}\'')
print('do NOT close any opening office application windows (minimize them instead)')

word = win32.gencache.EnsureDispatch('Word.Application')
word.DisplayAlerts = False
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.DisplayAlerts = False
ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
ppt.DisplayAlerts = constants.ppAlertsNone



def handle_error(path):
    logfile = open(pathlib.Path(str(current_dir) + 'log.txt'), 'w')
    print(f'ERROR: could not convert \'{path}\'')
    placeholder = open(path + '.txt', 'w')
    placeholder.write('file could not be converted')
    placeholder.close()
    relpath = path.replace(str(current_dir), '')
    newPath = 'C:\\FCI' + relpath
    os.makedirs(newPath.replace(os.path.basename(newPath), ''), exist_ok = True)
    shutil.copyfile(path, newPath)
    os.remove(path)
            
for path in pathlib.Path(str(current_dir)).rglob('*.*'):
    extension = pathlib.Path(path).suffix[1:].lower()

    path = str(path)
    #print(path)
    print(extension)
    if extension in ['doc', 'odt', 'docm']:
        print(path)
        try:
            doc = word.Documents.Open(path)
        except:
            handle_error(path)
            continue
        doc.Activate()

        if path.endswith('odt'):
            new_path = path[:-3] + 'docx'
        elif path.endswith('docm'):
            new_path = path[:-4] + 'docx'
        else:
            new_path = path + 'x'

        word.ActiveDocument.SaveAs(new_path, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(False)
        os.remove(path)
        
    elif extension in ['xls', 'xlsm', 'xlsb', 'ods']:
        print(path)      
        try:
            wb = excel.Workbooks.Open(path)
        except:
            handle_error(path)
            continue
            
        if path.endswith('ods'):
            new_path = path[:-3] + 'xlsx'
        elif path.endswith('xlsm'):
            new_path = path[:-4] + 'xlsx'
        elif path.endswith('xlsb'):
            new_path = path[:-4] + 'xlsx'
        else:
            new_path = path + 'x'
        
        wb.SaveAs(new_path, FileFormat = XLSX_FILE_FORMAT)
        wb.Close()
        os.remove(path)
        
    elif extension in ['ppt', 'pptm', 'odp']:
        print(path)
        try:
            presentation = ppt.Presentations.Open(path, WithWindow = False)
        except:
            handle_error(path)
            continue
        
        if path.endswith('odp'):
            new_path = path[:-3] + 'pptx'
        elif path.endswith('pptm'):
            new_path = path[:-4] + 'pptx'
        else:
            new_path = path + 'x'
            
        presentation.SaveAs(new_path, PPTX_FILE_FORMAT)
        presentation.Close()
        os.remove(path)
        
    elif extension in ['dot', 'dotm']:
        print(path) 
        print(f'problem with {extension}')
        # convert to dotx
        raise ValueError(f'extension not supported \'{extension}\'')
        
    elif extension in ['xlt', 'xltm']:
        print(path) 
        print(f'problem with {extension}')
        # convert to xltx
        raise ValueError(f'extension not supported \'{extension}\'')
        
    elif extension in ['pot', 'potm']:
        print(path) 
        print(f'problem with {extension}')
        # convert to potx
        raise ValueError(f'extension not supported \'{extension}\'')
        
    elif extension in ['pps', 'ppsm']:
        print(path)
        try:
            presentation = ppt.Presentations.Open(path, WithWindow = False)
        except:
            handle_error(path)
            continue
        
        if path.endswith('odp'):
            new_path = path[:-3] + 'pptx'
        elif path.endswith('pptm'):
            new_path = path[:-4] + 'pptx'
        else:
            new_path = path + 'x'
            
        presentation.SaveAs(new_path, PPTX_FILE_FORMAT)
        presentation.Close()
        os.remove(path)
       
try:     
    word.Application.Quit()
    excel.Application.Quit()
    ppt.Quit()
except:
    pass

input("Press Enter to continue...")