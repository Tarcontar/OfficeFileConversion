import os
import re
import sys
import pathlib
import binascii
import queue
import shutil
import win32com.client as win32
from win32com.client import constants
from time import sleep
import win32com
from multiprocessing import Process

#source = sys.argv[1]
#issue_target_dir = 'X:\\ZZ\\IF'
#legacy_target_dir = 'X:\\ZZ\\BF'
#logfile = open('X:\\ZZ\\log.txt', 'a')

source = 'C:\\Users\\admin\\Desktop\\OfficeFileConversion - Kopie\\source'
issue_target_dir = 'C:\\IF'
legacy_target_dir = 'C:\\BF'
logfile_path = 'C:\\log.txt'

DOCX_FILE_FORMAT = 12
DOTX_FILE_FORMAT = 14
PPTX_FILE_FORMAT = 24
POTX_FILE_FORMAT = 26
PPSX_FILE_FORMAT = 28
XLSX_FILE_FORMAT = 51
XLTX_FILE_FORMAT = 54
ACCDB_FILE_FORMAT = 54

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
    
try:
    access = win32.gencache.EnsureDispatch('Access.Application')
except AttributeError:
    shutil.rmtree(python_temp)
    access = win32.gencache.EnsureDispatch('Access.Application')

def get_magic(path):
    with open (path, 'rb') as myfile:
        header = myfile.read(4)
        return str(binascii.hexlify(header))


def copy_file(source, target):
    os.makedirs(target.replace(os.path.basename(target), ''), exist_ok = True)
    shutil.copyfile(source, target)


error_queue = queue.Queue()

def handle_errors():
    while True:
        logfile = None
        if not error_queue.empty():
            logfile = open(logfile_path, 'a')
        while not error_queue.empty():
            msg = error_queue.get()
            error_msg = f'ERROR: could not convert \'{msg}\' \n'
            print(error_msg)
            logfile.write(error_msg)
            placeholder = open(path + '.txt', 'w')
            placeholder.write('file could not be converted')
            placeholder.close()
            copy_file(path, issue_target_dir + path[2:])
            os.remove(path)
        if logfile is not None:
            logfile.close()
        sleep(1)
    
    
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
        error_queue.put(source)
        return False
    return True
    
    
def process_excel(source, target, format, target_dir):
    try:
        wb = excel.Workbooks.Open(source, UpdateLinks=False, Password='')
        wb.Application.DisplayAlerts = False
        wb.Application.EnableEvents = False
        wb.SaveAs(target, FileFormat=format, ConflictResolution=2)
        wb.Close()
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
    except Exception as e:
        print(e)
        error_queue.put(source)
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
        error_queue.put(source)
        
        
def process_access(source, target, format, target_dir):
    try:
        print(access)
        database = access.DBEngine.OpenDatabase(source)
        print(format)
        print(dir(access.DBEngine))
        #print(dir(database))
        #database.Activate()
        #access.ActiveDatabase.SaveAs(target, format)
        #database.SaveAs(target, format)
        #database.Close()
        #copy_file(source, target_dir + source[2:])
        #os.remove(source)
        print('done')
    except Exception as e:
        if hasattr(e, 'message'):
            print(e.message)
        print(e)
        #handle_error(source)


def process_file(path):
    path = str(path)
    print (path)
    
    if os.path.isdir(path):
        return 0

    extension = pathlib.Path(path).suffix[1:].lower()
    #print (os.path.basename(path))

    if os.path.basename(path).startswith('~$'):
        print (path)
        os.remove(path)
        return 0

    if extension in ['docx', 'doc', 'docm', 'dotx', 'dot', 'dotm', 'odt']:
        path, processing_needed = handle_fake_files(path, extension, ['docx', 'dotx'])
        if not processing_needed:
            return 0
        
        #print (path)
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
            
        #print (path)
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
        
        #print (path)
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
        
    #elif extension in ['accdb', 'mdb']: # accdt
    #    #print (path)
    #    format = ACCDB_FILE_FORMAT
    #    
    #    if extension in ['mdb']:
    #        new_path = path[:-3] + 'accdb'
    #    else:
    #        new_path = path
#
    #    for i in range(0, 100):
    #        process_access(path, path[:-6] + f'.te{i}', i, legacy_target_dir)
    #    return 1
        
    elif extension in ['msg', 'cmd', 'exe', 'msi', 'bat', 'lnk', 'reg', 'pol', 'ps1', 'psm1', 'psd1', 'ps1xml', 'pssc', 'psrc', 'cdxml']:
        #print (path)
        placeholder = open(path + '.txt', 'w')
        placeholder.write('file might be malicious and was moved to a backup location, please contact your IT')
        placeholder.close()
        copy_file(path, issue_target_dir + path[2:])
        os.remove(path)
    return 0
    
 
if __name__ == "__main__":
    print(f'Processing folder: {source}')
    file_count = 0
    issue_count = 0
    
    error_worker = Process(target=handle_errors)
    error_worker.start()
    
    for path in pathlib.Path(source).rglob('*.*'):
        try:
            file_count += process_file(path)
        except Exception as e:
            issue_count += 1
            path = str(path)
            if hasattr(e, 'message'):
                error_msg = f'ERROR: could not process \'{path}\' {e.message}\n'
            else:
                error_msg = f'ERROR: could not process \'{path}\'\n'
            print(error_msg)
            #logfile.write(error_msg)
            try:
                os.remove(path)
            except:
                pass

    error_worker.join()

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
    print(f'had {issue_count} issues')
    input('Press Enter to continue...')
