import os
import re
import sys
import pathlib
import binascii
import shutil
import win32com.client as win32
from win32com.client import constants
import win32com
from concurrent.futures import ProcessPoolExecutor

DOCX_FILE_FORMAT = 12
DOTX_FILE_FORMAT = 14
PPTX_FILE_FORMAT = 24
POTX_FILE_FORMAT = 26
PPSX_FILE_FORMAT = 28
XLSX_FILE_FORMAT = 51
XLTX_FILE_FORMAT = 54

X_FILE_MAGIC = '504b0304'

current_dir = pathlib.Path(__file__).parent.absolute()
print(f'processing all [\'doc\', \'docm\', \'odt\', \'xls\', \'xlsm\', \'xlsb\', \'ods\', \'ppt\', \'pptm\', \'odp\'] files in \'{current_dir}\'')
print('do NOT close any opening office application windows (minimize them instead)')

source = 'C:\\Users\\admin\\Desktop\\OfficeFileConversion\\source'
issue_target_dir = 'C:\\ZZ\\IF'
legacy_target_dir = 'C:\\ZZ\\BF'

logfile = open('C:\\ZZ\\log.txt', 'a')

word = win32.gencache.EnsureDispatch('Word.Application')
word.DisplayAlerts = False
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.DisplayAlerts = False
excel.EnableEvents = False
ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
ppt.DisplayAlerts = constants.ppAlertsNone


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
    copy_file(path, issue_target_dir + path[2:])
    os.remove(path)
    
    
def handle_fake_files(path, extension, extension_filter):
    if not extension in [extension_filter]:
        return path, True
    if X_FILE_MAGIC in get_magic(path):
        return path, False
    print('WARNING: fake file detected')
    os.rename(path, path[:-1])
    path = path[:-1]
    return path, True
    

def process_file(path):
    if os.path.isdir(path):
        return 0

    extension = pathlib.Path(path).suffix[1:].lower()

    path = str(path)
    #print(os.path.basename(path))
    #if extension in ['docxx']:
    #    os.rename(path, path[:-1])
    #    extension = 'docx'
    #    path = path[:-1]
        
    #if extension in ['xlsxx']:
    #    os.rename(path, path[:-1])
    #    extension = 'xlsx'
    #    path = path[:-1]
        
    if extension in ['docx', 'doc', 'docm', 'dot', 'dotm', 'odt']:
        path, processing_needed = handle_fake_files(path, extension, 'docx')
        if not processing_needed:
            return 0
        print(path)
        
        try:
            doc = word.Documents.Open(path, ConfirmConversions=False, Visible=False)
            doc.Activate()
        except Exception as e:
            print(e)
            handle_error(path)
            return 0
        
        if extension in ['dot', 'dotm']:
            ff = DOTX_FILE_FORMAT
        else:
            ff = DOCX_FILE_FORMAT

        if extension in ['odt']:
            new_path = path[:-3] + 'docx'
        elif extension in ['docm']:
            new_path = path[:-1] + 'x'
        elif extension in ['dotm']:
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'

        word.ActiveDocument.SaveAs(new_path, ff)
        doc.Close(False)
        copy_file(path, legacy_target_dir + path[2:])
        os.remove(path)
        return 1
        
    elif extension in ['xlsx', 'xls', 'xlsm', 'xlsb', 'xlt', 'xltm', 'ods']:
        path, processing_needed = handle_fake_files(path, extension, 'xlsx')
        if not processing_needed:
            return 0
        print(path)
        
        try:
            wb = excel.Workbooks.Open(path, Password='')
            wb.Application.DisplayAlerts = False
        except Exception as e:
            print(e)
            handle_error(path)
            return 0
            
        if extension in ['xlt', 'xltm']:
            ff = XLTX_FILE_FORMAT
        else:
            ff = XLSX_FILE_FORMAT
            
        if extension in ['ods']:
            new_path = path[:-3] + 'xlsx'
        elif extension in ['xlsm']:
            new_path = path[:-1] + 'x'
        elif extension in ['xlsb']:
            new_path = path[:-1] + 'x'
        elif extension in ['xltm']:
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'
        
        wb.SaveAs(new_path, FileFormat=ff, ConflictResolution=2)
        wb.Close()
        copy_file(path, legacy_target_dir + path[2:])
        os.remove(path)
        return 1
        
    elif extension in ['pptx', 'ppt', 'pptm', 'pot', 'potm', 'pps', 'ppsm', 'odp']:
        path, processing_needed = handle_fake_files(path, extension, 'pptx')
        if not processing_needed:
            return 0
        print(path)
        
        try:
            presentation = ppt.Presentations.Open(path, WithWindow=False)
        except Exception as e:
            print(e)
            handle_error(path)
            return 0
            
        if extension in ['pot', 'potm']:
            ff = POTX_FILE_FORMAT
        elif extension in ['pps', 'ppsm']:
            ff = PPSX_FILE_FORMAT
        else:
            ff = PPTX_FILE_FORMAT
        
        if extension in ['odp']:
            new_path = path[:-3] + 'pptx'
        elif extension in ['pptm']:
            new_path = path[:-1] + 'x'
        elif extension in ['potm']:
            new_path = path[:-1] + 'x'
        elif extension in ['ppsm']:
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'
            
        presentation.SaveAs(new_path, ff)
        presentation.Close()
        copy_file(path, legacy_target_dir + path[2:])
        os.remove(path)
        return 1
        
    elif extension in ['exe', 'msi', 'bat', 'lnk', 'reg', 'pol', 'ps1', 'psm1', 'psd1', 'ps1xml', 'pssc', 'psrc', 'cdxml']:
        copy_file(path, legacy_target_dir + path[2:])
        os.remove(path)
    return 0
    
 
def main():
    print(f'Processing folder: {source}')
    file_count = 0
    for path in pathlib.Path(source).rglob('*.*'):
        try:
            print (str(path))
            file_count += process_file(path)
        except Exception as e:
            path = str(path)
            if hasattr(e, 'message'):
                error_msg = f'ERROR: could not process \'{path}\' {e.message}\n'
            else:
                error_msg = f'ERROR: could not process \'{path}\'\n'
            print(error_msg)
            logfile.write(error_msg)
            os.remove(path)

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
    print(f'converted {file_count} files')
    input('Press Enter to continue...')


if __name__ == "__main__":
    main()