import os
import sys
import pathlib
import binascii
import shutil
import pythoncom
import zipfile
import win32com.client as win32
from win32com.client import constants
import win32com
from subprocess import Popen


process_malicious = True #if len(sys.argv) >= 3 and sys.argv[2] in ['True', 'true'] else False

target_dir = sys.argv[1]
issue_target_dir = f'{target_dir}\\IF'
legacy_target_dir = f'{target_dir}\\BF'
logfile = open(f'{target_dir}\\log.txt', 'a')

source = sys.argv[2]

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
    if not os.path.exists(target):
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
    

def process_word(word, source, target, format):
    try:
        if word is None:
            word = setup_word()
        doc = word.Documents.Open(source, ConfirmConversions=False, Visible=False, PasswordDocument="invalid")
        doc.Activate()
        word.ActiveDocument.SaveAs(target, format)
        doc.Close(False)
        os.remove(source)
        return
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        if error.args[0] == -2147352567:
            print('-> file is password protected')
        else:
            print(error)
    except Exception as e:
        print(e)
        
    handle_error(source)
    print('Exception occured with word')
    
    
def process_excel(excel, source, target, format):
    try:
        if excel is None:
            excel = setup_excel()
        wb = excel.Workbooks.Open(source, UpdateLinks=False, Password='', WriteResPassword='')
        wb.Application.DisplayAlerts = False
        wb.Application.EnableEvents = False
        wb.SaveAs(target, FileFormat=format, ConflictResolution=2)
        wb.Close()
        os.remove(source)
        return
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        if error.args[0] == -2147352567:
            print('-> file is password protected')
        else:
            print(error)
    except Exception as e:
        print(e)
        
    print('Exception occured in excel')
    handle_error(source)
    
    
def process_powerpoint(ppt, source, target, format):
    try:
        if ppt is None:
            ppt = setup_ppt()
        presentation = ppt.Presentations.Open(source + ':::', WithWindow=False)
        presentation.SaveAs(target, format)
        presentation.Close()
        os.remove(source)
        return
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        if error.args[0] == -2147352567:
            print('-> file is password protected')
        else:
            print(error)
    except Exception as e:
        print(e)

    print('Exception occured in powerpoint')
    handle_error(source)
        
        
def process_outlook(word, excel, ppt, outlook, source):
    try:
        if outlook is None:
            outlook = setup_outlook()
 
        tmp_file = issue_target_dir + source[2:]
        copy_file(source, tmp_file) # TODO: only workaround for outlook not closing file properly
        os.remove(source)
        msg = outlook.OpenSharedItem(tmp_file)
        
        html_path = source[:-4] + '.html'
        msg.SaveAs(html_path, constants.olHTML)
        doc = word.Documents.Open(html_path)
        doc.ExportAsFixedFormat(source[:-4] + '.pdf', 17)
        doc.Close(False)
        os.remove(html_path)
        
        try:
            shutil.rmtree(source[:-4] + '_files')
        except:
            shutil.rmtree(source[:-4] + '-Dateien')
                
        if not msg.Attachments:
            return
        
        count = 0
        for attachment in msg.Attachments:
            path = source[:-3] + attachment.FileName
            print(attachment.FileName)
            attachment.SaveAsFile(path)
            count += process_file(word, excel, ppt, outlook, path)
        
        msg.Close(1)  
        return
            
    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return
        print(e)
    except pythoncom.com_error as error:
        print(error)
    except Exception as e:
        print(e)
        
    print('Exception occured in outlook')
    handle_error(source) 
        

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
        print('## extracting...')
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
        print('## compressing...')
        #process = Popen(['C:\\Program Files\\7-Zip\\7z.exe', 'a', '-mmt=24', target_path + '.zip', target_path + '\\*'])
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
        
        process_word(word, path, new_path, format)
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
        
        process_excel(excel, path, new_path, format)
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

        process_powerpoint(ppt, path, new_path, format)
        return 1
        
    elif extension in outlook_filter:
        process_outlook(word, excel, ppt, outlook, path)
        return 1
        
    elif process_malicious and extension in malicious_filter:
        #print (path)
        if 'Win-Plantafel2' in path:
            return 0
        
        placeholder_path = path + '.txt'
        if not os.path.exists(placeholder_path):
            placeholder = open(placeholder_path, 'w')
            placeholder.write('file might be malicious and was moved to a backup location, please contact your IT')
            placeholder.close()
        copy_file(path, issue_target_dir + path[2:])
        os.remove(path)
        
    elif extension in archive_filter and zipfile.is_zipfile(path):
        return process_zip(word, excel, ppt, outlook, path)
    return 0
    
    
def setup_office_app(func):
    try:
        return func()
    except AttributeError:
        shutil.rmtree(python_temp)
        return func()
      
      
def shutdown_office_app(app):
    try:     
        app.Quit()
    except:
        pass
 
 
if __name__ == "__main__":
    print(f'Processing folder: {source}')
    file_count = 0
    issue_count = 0
    
    word = setup_office_app(setup_word)
    excel = setup_office_app(setup_excel)
    ppt = setup_office_app(setup_ppt)
    outlook = setup_office_app(setup_outlook)
    
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
            #input()
            try:
                os.remove(path)
            except:
                pass

    shutdown_office_app(word.Application)
    shutdown_office_app(excel.Application)
    shutdown_office_app(ppt)
    shutdown_office_app(outlook)

    logfile.close()
    print(f'converted {file_count} files')
    print(f'had {issue_count} issues')
    input('Press Enter to continue...')
