import os
import re
import sys
import pathlib
import binascii
import shutil
import win32com.client as win32
from win32com.client import constants
import win32com
from multiprocessing import Process

source = sys.argv[1]
issue_target_dir = 'X:\\ZZ\\IF'
legacy_target_dir = 'X:\\ZZ\\BF'
logfile = open('X:\\ZZ\\log.txt', 'a')

DOCX_FILE_FORMAT = 12
DOTX_FILE_FORMAT = 14
PPTX_FILE_FORMAT = 24
POTX_FILE_FORMAT = 26
PPSX_FILE_FORMAT = 28
XLSX_FILE_FORMAT = 51
XLTX_FILE_FORMAT = 54

ZIP_FILE_MAGIC = '504b0304'

current_dir = pathlib.Path(__file__).parent.absolute()
print(f'processing all [docx, doc, docm, dot, dotm, odt, xlsx, xls, xlsm, xlsb, xlt, xltm, ods, pptx, ppt, pptm, pot, potm, pps, ppsm, odp] files in \'{source}\'')
print('do NOT close any opening office application windows (minimize them instead)')

python_temp = 'C:\\Users\\admin\\AppData\\Local\\Temp\\gen_py'


try:
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.DisplayAlerts = False
except AttributeError:
    shutil.rmtree(python_temp)
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.DisplayAlerts = False
    
try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.EnableEvents = False
except AttributeError:
    shutil.rmtree(python_temp)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    
try:
    ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
    ppt.DisplayAlerts = constants.ppAlertsNone
except AttributeError:
    shutil.rmtree(python_temp)
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
    
    
def handle_fake_files(path, extension, extensions_filter):
    if not extension in extensions_filter:
        return path, True
    if ZIP_FILE_MAGIC in get_magic(path):
        return path, False
    print('WARNING: fake file detected')
    os.rename(path, path[:-1])
    path = path[:-1]
    return path, True
    

def process_word(source, target, format, target_dir):
    try:
        doc = word.Documents.Open(source, ConfirmConversions=False, Visible=False, PasswordDocument="invalid")
        doc.Activate()
        word.ActiveDocument.SaveAs(target, format)
        doc.Close(False)
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
    except Exception as e:
        print(e)
        handle_error(source)
        return False
    return True
    
    
def process_excel(source, target, format, target_dir):
    try:
        wb = excel.Workbooks.Open(source, Password='')
        wb.Application.DisplayAlerts = False
        wb.SaveAs(target, FileFormat=format, ConflictResolution=2)
        wb.Close()
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
    except Exception as e:
        print(e)
        handle_error(source)
        return False
    return True
    
    
def process_powerpoint(source, target, format, target_dir):
    try:
        presentation = ppt.Presentations.Open(source + ':::', WithWindow=False)
        presentation.SaveAs(target, format)
        presentation.Close()
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
    except Exception as e:
        print(e)
        handle_error(source)


def process_file(path):
    if os.path.isdir(path):
        return 0

    extension = pathlib.Path(path).suffix[1:].lower()
    path = str(path)

    #print (path)
    #print (os.path.basename(path))
    if os.path.basename(path).startswith('~$'):
        print (path)
        os.remove(path)
        return 0
        
    if extension in ['docx', 'doc', 'docm', 'dotx', 'dot', 'dotm', 'odt']:
        path, processing_needed = handle_fake_files(path, extension, ['docx', 'dotx'])
        if not processing_needed:
            return 0
        
        print (path)
        if extension in ['dotx', 'dot', 'dotm']:
            format = DOTX_FILE_FORMAT
        else:
            format = DOCX_FILE_FORMAT

        if extension in ['odt']:
            new_path = path[:-3] + 'docx'
        elif extension in ['docm', 'dotm']:
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'
        
        process_word(path, new_path, format, legacy_target_dir)
        return 1
        
    elif extension in ['xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xlt', 'xltm', 'ods']:
        path, processing_needed = handle_fake_files(path, extension, ['xlsx', 'xltx'])
        if not processing_needed:
            return 0
            
        print (path)
        if extension in ['xltx', 'xlt', 'xltm']:
            format = XLTX_FILE_FORMAT
        else:
            format = XLSX_FILE_FORMAT
            
        if extension in ['ods']:
            new_path = path[:-3] + 'xlsx'
        elif extension in ['xlsm', 'xlsb', 'xltm']:
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'
        
        process_excel(path, new_path, format, legacy_target_dir)
        return 1

    elif extension in ['pptx', 'ppt', 'pptm', 'potx', 'pot', 'potm', 'ppsx', 'pps', 'ppsm', 'odp']:
        path, processing_needed = handle_fake_files(path, extension, ['pptx', 'potx', 'ppsx'])
        if not processing_needed:
            return 0
        
        print (path)
        if extension in ['potx', 'pot', 'potm']:
            format = POTX_FILE_FORMAT
        elif extension in ['ppsx', 'pps', 'ppsm']:
            format = PPSX_FILE_FORMAT
        else:
            format = PPTX_FILE_FORMAT
        
        if extension in ['odp']:
            new_path = path[:-3] + 'pptx'
        elif extension in ['pptm', 'potm', 'ppsm']:
            new_path = path[:-1] + 'x'
        else:
            new_path = path + 'x'

        process_powerpoint(path, new_path, format, legacy_target_dir)
        return 1
        
    elif extension in ['msg', 'exe', 'msi', 'bat', 'lnk', 'reg', 'pol', 'ps1', 'psm1', 'psd1', 'ps1xml', 'pssc', 'psrc', 'cdxml']:
        print (path)
        placeholder = open(path + '.txt', 'w')
        placeholder.write('file might be malicious and was moved to a backup location, please contact your IT')
        placeholder.close()
        copy_file(path, issue_target_dir + path[2:])
        os.remove(path)
    return 0
    
 
if __name__ == "__main__":
    print(f'Processing folder: {source}')
    file_count = 0
    for path in pathlib.Path(source).rglob('*.*'):
        try:
            file_count += process_file(path)
        except Exception as e:
            path = str(path)
            if hasattr(e, 'message'):
                error_msg = f'ERROR: could not process \'{path}\' {e.message}\n'
            else:
                error_msg = f'ERROR: could not process \'{path}\'\n'
            print(error_msg)
            logfile.write(error_msg)
            try:
                os.remove(path)
            except:
                pass

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
