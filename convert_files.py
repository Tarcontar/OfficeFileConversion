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

issue_target_dir = ''
legacy_target_dir = ''
logfile = None

source = ''

python_temp = 'C:\\Users\\admin\\AppData\\Local\\Temp\\gen_py'

word_filter = ['doc', 'docm', 'dot', 'dotm', 'odt']
word_fake_filter = ['docx', 'dotx']
excel_filter = ['xls', 'xlsm', 'xlsb', 'xlt', 'xltm', 'ods']
excel_fake_filter = ['xlsx', 'xltx']
ppt_filter = ['ppt', 'pptm', 'pot', 'potm', 'pps', 'ppsm', 'odp']
ppt_fake_filter = ['pptx', 'potx', 'ppsx']
outlook_filter = ['msg']
archive_filter = ['zip', 'rar', '7z']
malicious_filter = ['pst', 'xlam', 'osd', 'py', 'exe', 'msi', 'bat', 'reg', 'pol', 'ps1', 'psm1', 'psd1', 'ps1xml', 'pssc', 'psrc', 'cdxml']


 #TODO: get from sys.argv
process_word = True
process_excel = True
process_ppt = True
process_outlook = True
process_fakefiles = True
process_malicious = True
process_archives = True

word = None
excel = None
ppt = None
outlook = None


def setup_word():
    global word
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.DisplayAlerts = False
    
    
def setup_excel():
    global excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.EnableEvents = False


def setup_ppt():
    global ppt
    ppt = win32.gencache.EnsureDispatch('Powerpoint.Application')
    ppt.DisplayAlerts = constants.ppAlertsNone
    
    
def setup_outlook():
    global outlook
    outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')


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
    print('## copying...')
    copy_file(path, issue_target_dir + path[2:])
    print('## deleting...')
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
    

def process_word(source, target, format):
    try:
        if word is None:   
            setup_word()
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
    
    
def process_excel(source, target, format):
    try:
        if excel is None:
            setup_excel()
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
    
    
def process_powerpoint(source, target, format):
    try:
        if ppt is None:
            setup_ppt()
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
        
        
def process_outlook(source):
    try:
        if outlook is None:
            setup_outlook()
 
        tmp_file = legacy_target_dir + source[2:]
        copy_file(source, tmp_file) # TODO: only workaround for outlook not closing file properly
        os.remove(source)
        msg = outlook.OpenSharedItem(tmp_file)
        
        html_path = source[:-4] + '.html'
        msg.SaveAs(html_path, constants.olHTML)
        doc = word.Documents.Open(html_path)
        doc.ExportAsFixedFormat(source[:-4] + '.pdf', 17)
        doc.Close(False)
        os.remove(html_path)
        
        if os.path.exists(source[:-4] + '_files'):
            shutil.rmtree(source[:-4] + '_files')
        if os.path.exists(source[:-4] + '-Dateien'):
            shutil.rmtree(source[:-4] + '-Dateien')
                
        if not msg.Attachments:
            return
        
        count = 0
        for attachment in msg.Attachments:
            path = source[:-3] + attachment.FileName
            attachment.SaveAsFile(path)
            count += process_file(path)
        
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
        

def process_zip(source):
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
                needs_processing = True
                break
            if process_fakefiles and (extension in word_fake_filter \
                        or extension in excel_fake_filter \
                        or extension in ppt_fake_filter):
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
                count += process_file(path)
            except Exception as e:
                handle_error(source)
                shutil.rmtree(target_path)
                return 0
                  
        os.remove(source)
        print('## compressing...')
        shutil.make_archive(target_path, 'zip', target_path)
        shutil.rmtree(target_path)
        return count

    except WindowsError as e:
        if e.winerror == ACCESS_DENIED:
            return 0
        print(e)
        
    except Exception as e:
        print(e)
        handle_error(source)
        return 0


def process_file(path):
    path = str(path)
    print (path)
    
    if os.path.isdir(path):
        return 0

    if os.path.basename(path).startswith('~$'):
        os.remove(path)
        return 0
        
    extension = pathlib.Path(path).suffix[1:].lower()
    #print (os.path.basename(path))
    
    if process_word and (extension in word_filter or process_fakefiles and extension in word_fake_filter):
        path, processing_needed = handle_fake_files(path, extension, word_fake_filter)
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
        
        process_word(path, new_path, format)
        return 1
        
    elif process_excel and (extension in excel_filter or process_fakefiles and extension in excel_fake_filter):
        path, processing_needed = handle_fake_files(path, extension, excel_fake_filter)
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
        
        process_excel(path, new_path, format)
        return 1

    elif process_ppt and (extension in ppt_filter or process_fakefiles and extension in ppt_fake_filter):
        path, processing_needed = handle_fake_files(path, extension, ppt_fake_filter)
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

        process_powerpoint(path, new_path, format)
        return 1
        
    elif process_outlook and extension in outlook_filter:
        process_outlook(path)
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
        
    elif process_archives and extension in archive_filter:
        return process_zip(word, excel, ppt, outlook, path)
    return 0
    
    
def setup_office_app(func):
    try:
        func()
    except AttributeError:
        shutil.rmtree(python_temp)
        func()
      
      
def shutdown_office_app(app):
    try:     
        app.Quit()
    except:
        pass
 
 
def process_folder(target_dir, source):
    global issue_target_dir
    issue_target_dir = f'{target_dir}\\IF'
    global legacy_target_dir
    legacy_target_dir = f'{target_dir}\\BF'
    
    logfile_path = f'{target_dir}\\log.txt'
    
    global logfile
    if os.path.exists(logfile_path):
        logfile = open(logfile_path, 'a')
    else:
        logfile = open(logfile_path, 'w')

    print(f'processing all {word_filter} files in \'{source}\'')
    print(f'processing all {excel_filter} files in \'{source}\'')
    print(f'processing all {ppt_filter} files in \'{source}\'')
    print(f'processing all {outlook_filter} files in \'{source}\'')
    print(f'processing all {malicious_filter} files in \'{source}\'')
    print('do NOT close any opening office application windows (minimize them instead)')
 
    file_count = 0
    issue_count = 0
    
    setup_office_app(setup_word)
    setup_office_app(setup_excel)
    setup_office_app(setup_ppt)
    setup_office_app(setup_outlook)
    
    if os.path.isdir(source):
        paths = pathlib.Path(source).rglob('*.*')
    else:
        paths = [source]
    
    for path in paths:
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
            logfile.write(error_msg)
            print('press any key to continue...')
            #input()
            try:
                os.remove(path)
            except:
                pass
        except KeyboardInterrupt:
            break

    shutdown_office_app(word.Application)
    shutdown_office_app(excel.Application)
    shutdown_office_app(ppt)
    shutdown_office_app(outlook)

    logfile.close()
    print(f'converted {file_count} files')
    print(f'had {issue_count} issues')
    return (file_count, issue_count)
    
    
    
if __name__ == "__main__":
    process_folder(sys.argv[1], sys.argv[2])
    input('Press Enter to continue...')

