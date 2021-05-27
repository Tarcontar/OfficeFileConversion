import os
import re
import sys
import pathlib
import binascii
import shutil
import pythoncom
import zipfile
import win32com.client as win32
from win32com.client import constants
import win32com


process_malicious = True #if len(sys.argv) >= 3 and sys.argv[2] in ['True', 'true'] else False

source = sys.argv[1]
issue_target_dir = 'X:\\ZZ\\IF'
legacy_target_dir = 'X:\\ZZ\\BF'
logfile = open('X:\\ZZ\\log.txt', 'a')

ACCESS_DENIED = 5


DOCX_FILE_FORMAT = 12
DOTX_FILE_FORMAT = 14
PPTX_FILE_FORMAT = 24
POTX_FILE_FORMAT = 26
PPSX_FILE_FORMAT = 28
XLSX_FILE_FORMAT = 51
XLTX_FILE_FORMAT = 54

ZIP_FILE_MAGIC = '504b0304'
#EXE_FILE_MAGICS = ['4d5a', '5a4d']

current_dir = pathlib.Path(__file__).parent.absolute()
python_temp = 'C:\\Users\\admin\\AppData\\Local\\Temp\\gen_py'

word_filter = ['docx', 'doc', 'docm', 'dotx', 'dot', 'dotm', 'odt']
excel_filter = ['xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xlt', 'xltm', 'ods']
ppt_filter = ['pptx', 'ppt', 'pptm', 'potx', 'pot', 'potm', 'ppsx', 'pps', 'ppsm', 'odp']
outlook_filter = ['msg']
archive_filter = ['zip', 'rar', '7z']
malicious_filter = ['pst', 'xlam', 'osd', 'py', 'exe', 'msi', 'bat', 'reg', 'pol', 'ps1', 'psm1', 'psd1', 'ps1xml', 'pssc', 'psrc', 'cdxml']

print(f'processing all {word_filter} files in \'{source}\'')
print(f'processing all {excel_filter} files in \'{source}\'')
print(f'processing all {ppt_filter} files in \'{source}\'')
print(f'processing all {outlook_filter} files in \'{source}\'')
print(f'processing all {malicious_filter} files in \'{source}\'')
print('do NOT close any opening office application windows (minimize them instead)')



def setup_word():
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.DisplayAlerts = False
    return word
    
    
def setup_excel():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    return excel


def setup_ppt():
    ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
    ppt.DisplayAlerts = constants.ppAlertsNone
    return ppt
    
    
def setup_outlook():
    outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    return outlook


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
    placeholder.write('file could not be converted please contact your IT')
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
    

def process_word(word, source, target, format, target_dir):
    try:
        if word is None:
            word = setup_word()
        doc = word.Documents.Open(source, ConfirmConversions=False, Visible=False, PasswordDocument="invalid")
        doc.Activate()
        word.ActiveDocument.SaveAs(target, format)
        doc.Close(False)
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
        return
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        print(error)
    except Exception as e:
        print(e)
        
    handle_error(source)
    print('Exception occured with word')
    
    
def process_excel(excel, source, target, format, target_dir):
    try:
        if excel is None:
            excel = setup_excel()
        wb = excel.Workbooks.Open(source, UpdateLinks=False, Password='', WriteResPassword='')
        wb.Application.DisplayAlerts = False
        wb.Application.EnableEvents = False
        wb.SaveAs(target, FileFormat=format, ConflictResolution=2)
        wb.Close()
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
        return
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        print(error)
    except Exception as e:
        print(e)
        
    print('Exception occured in excel')
    handle_error(source)
    
    
def process_powerpoint(ppt, source, target, format, target_dir):
    try:
        if ppt is None:
            ppt = setup_ppt()
        presentation = ppt.Presentations.Open(source + ':::', WithWindow=False)
        presentation.SaveAs(target, format)
        presentation.Close()
        #copy_file(source, target_dir + source[2:])
        os.remove(source)
        return
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        print(error)
    except Exception as e:
        print(e)

    print('Exception occured in powerpoint')
    handle_error(source)
        
        
def process_outlook(word, excel, ppt, outlook, source):
    try:
        if outlook is None:
            outlook = setup_outlook()
        msg = outlook.OpenSharedItem(source)
        print(source)
        #msg.ExportAsFixedFormat(2, source[:-4])
        msg.PrintOut()
        
        if not msg.Attachments:
            os.remove(source)
            return
        
        directory = source[:-4]
        try:
            os.mkdir(directory)
        except:
            pass
        
        for attachment in msg.Attachments:
            attachment.SaveAsFile(directory + '\\' + attachment.FileName)
             
        count = 0
        for path in pathlib.Path(directory).rglob('*.*'):
            count += process_file(word, excel, ppt, outlook, path)
                
        msg.Close(1)
        os.remove(source)
        shutil.rmtree(directory)
        return
            
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            print(e)
            return
        print(e)
    except pythoncom.com_error as error:
        print(error)
    except Exception as e:
        print(e)
        
    print('Exception occured in outlook')
    handle_error(source) 
    shutil.rmtree(directory)
        

def process_zip(word, excel, ppt, outlook, source):
    try:
        zip = zipfile.ZipFile(source)
        for zinfo in zip.infolist():
            is_encrypted = zinfo.flag_bits & 0x1 
            if is_encrypted:
                print (f'WARNING: {source} is encrypted!')
                zip.close()
                handle_error(source)
                return 0
                
        needs_processing = False
        for file in zip.namelist():
            extension = pathlib.Path(file).suffix[1:].lower()
            
            if extension in word_filter or extension in excel_filter \
                or extension in ppt_filter or extension in outlook_filter \
                    or extension in malicious_filter or extension in archive_filter:
                if not extension in ['docx', 'dotx', 'xlsx', 'xltx', 'pptx', 'potx', 'ppsx']:
                    needs_processing = True
                    break
            
        if not needs_processing:
            return 1
            
        target_path = source[:-4]
        zip.extractall(target_path)
        zip.close()
            
        count = 0
        for path in pathlib.Path(target_path).rglob('*.*'):
            try:
                count += process_file(word, excel, ppt, outlook, path)
            except Exception as e:
                handle_error(source)
                shutil.rmtree(target_path)
                return
                  
        os.remove(source)
        shutil.make_archive(target_path, 'zip', target_path)
        shutil.rmtree(target_path)
        return count

    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
        
    except Exception as e:
        print(e)
        handle_error(source)
        return 0


def process_file(word, excel, ppt, outlook, path):
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
    
    if extension in word_filter:
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
        
        process_word(word, path, new_path, format, legacy_target_dir)
        return 1
        
    elif extension in excel_filter:
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
        
        process_excel(excel, path, new_path, format, legacy_target_dir)
        return 1

    elif extension in ppt_filter:
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

        process_powerpoint(ppt, path, new_path, format, legacy_target_dir)
        return 1
        
    elif extension in outlook_filter:
        process_outlook(word, excel, ppt, outlook, path)
        return 1
        
    elif process_malicious and extension in malicious_filter:
        #print (path)
        if 'Win-Plantafel2' in path:
            return 0
        
        placeholder = open(path + '.txt', 'w')
        placeholder.write('file might be malicious and was moved to a backup location, please contact your IT')
        placeholder.close()
        copy_file(path, issue_target_dir + path[2:])
        os.remove(path)
        
    elif extension in archive_filter and zipfile.is_zipfile(path):
        return process_zip(word, excel, ppt, outlook, path)
    return 0
    
 
if __name__ == "__main__":
    print(f'Processing folder: {source}')
    file_count = 0
    issue_count = 0
    
    try:
        word = setup_word()
    except AttributeError:
        shutil.rmtree(python_temp)
        word = setup_word()
    
    try:
        excel = setup_excel()
    except AttributeError:
        shutil.rmtree(python_temp)
        excel = setup_excel()
        
    try:
        ppt = setup_ppt()
    except AttributeError:
        shutil.rmtree(python_temp)
        ppt = setup_ppt()
        
    try:
        outlook = setup_outlook()
    except AttributeError:
        shutil.rmtree(python_temp)
        outlook = setup_outlook()
    
    for path in pathlib.Path(source).rglob('*.*'):
        try:
            file_count += process_file(word, excel, ppt, outlook, path)
        except Exception as e:
            issue_count += 1
            path = str(path)
            if hasattr(e, 'message'):
                error_msg = f'ERROR: could not process \'{path}\' {e.message}\n'
            else:
                error_msg = f'ERROR: could not process \'{path}\'\n'
            print(error_msg)
            logfile.write(error_msg)
            print('press any key to continue...')
            input()
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
        
    try:     
        outlook.Quit()
    except:
        pass

    logfile.close()
    print(f'converted {file_count} files')
    print(f'had {issue_count} issues')
    input('Press Enter to continue...')
