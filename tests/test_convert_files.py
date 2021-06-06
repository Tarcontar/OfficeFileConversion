import os
import sys
import unittest
import shutil
import zipfile
from os.path import dirname as up
from utils import TestCase

sys.path.insert(1, '..')
from convert_files import process_folder, get_magic, ZIP_FILE_MAGIC


class FileConversionTests(TestCase):
    def copy_file(self, filename):
        source_file = up(self.relpath()) + os.path.sep +  'source' + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
    def run_test(self, filename, expected, result=(1, 0), check_file=True, check_magic=True):
        self.copy_file(filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual(result, process_folder(issue_dir, target_dir))
        if check_file:
            self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        if check_magic:
            self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))

    def test_convert_doc(self):
        filename = 'doc.doc'
        expected = 'doc.docx'
        self.run_test(filename, expected)
        
    def test_convert_docm(self):
        filename = 'docm.docm'
        expected = 'docm.docx'
        self.run_test(filename, expected)
        
    def test_convert_dot(self):
        filename = 'dot.dot'
        expected = 'dot.dotx'
        self.run_test(filename, expected)
        
    def test_convert_docx(self):
        filename = 'docx.docx'
        expected = 'docx.docx'
        self.run_test(filename, expected, result=(0, 0), check_file=False)

    def test_convert_doc_fake_docx(self):
        filename = 'doc_fake_docx.docx'
        expected = 'doc_fake_docx.docx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_dot_fake_dotx(self):
        filename = 'dot_fake_dotx.dotx'
        expected = 'dot_fake_dotx.dotx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_dotm(self):
        filename = 'dotm.dotm'
        expected = 'dotm.dotx'
        self.run_test(filename, expected)
        
    def test_convert_odt(self):
        filename = 'odt.odt'
        expected = 'odt.docx'
        self.run_test(filename, expected)
        
    def test_convert_doc_password(self):
        filename = 'doc_password.doc'
        expected = 'doc_password.doc.txt'
        self.run_test(filename, expected, result=(0, 1), check_magic=False)
        
    def test_convert_autosave(self):
        filename = '~$autosave.docx'
        self.copy_file(filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((0, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
       
    def test_convert_xls(self):
        filename = 'xls.xls'
        expected = 'xls.xlsx'
        self.run_test(filename, expected)
        
    def test_convert_xls_fake_xlsx(self):
        filename = 'xls_fake_xlsx.xlsx'
        expected = 'xls_fake_xlsx.xlsx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_xls_password(self):
        filename = 'xls_password.xls'
        expected = 'xls_password.xls.txt'
        self.run_test(filename, expected, result=(0, 1), check_magic=False)
        
    def test_convert_xlsb(self):
        filename = 'xlsb.xlsb'
        expected = 'xlsb.xlsx'
        self.run_test(filename, expected)
      
    def test_convert_xlsm(self):
        filename = 'xlsm.xlsm'
        expected = 'xlsm.xlsx'
        self.run_test(filename, expected)
        
    def test_convert_xlt(self):
        filename = 'xlt.xlt'
        expected = 'xlt.xltx'
        self.run_test(filename, expected)
        
    def test_convert_xlt_fake_xltx(self):
        filename = 'xlt_fake_xltx.xltx'
        expected = 'xlt_fake_xltx.xltx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_xltm(self):
        filename = 'xltm.xltm'
        expected = 'xltm.xltx'
        self.run_test(filename, expected)
        
    def test_convert_ods(self):
        filename = 'ods.ods'
        expected = 'ods.xlsx'
        self.run_test(filename, expected)
        
    def test_convert_pot(self):
        filename = 'pot.pot'
        expected = 'pot.potx'
        self.run_test(filename, expected)
        
    def test_convert_pot_fake_potx(self):
        filename = 'pot_fake_potx.potx'
        expected = 'pot_fake_potx.potx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_potm(self):
        filename = 'potm.potm'
        expected = 'potm.potx'
        self.run_test(filename, expected)
        
    def test_convert_pps(self):
        filename = 'pps.pps'
        expected = 'pps.ppsx'
        self.run_test(filename, expected)
        
    def test_convert_pps_fake_ppsx(self):
        filename = 'pps_fake_ppsx.ppsx'
        expected = 'pps_fake_ppsx.ppsx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_ppsm(self):
        filename = 'ppsm.ppsm'
        expected = 'ppsm.ppsx'
        self.run_test(filename, expected)
       
    def test_convert_ppt(self):
        filename = 'ppt.ppt'
        expected = 'ppt.pptx'
        self.run_test(filename, expected)
        
    def test_convert_ppt_fake_pptx(self):
        filename = 'ppt_fake_pptx.pptx'
        expected = 'ppt_fake_pptx.pptx'
        self.run_test(filename, expected, check_file=False)
        
    def test_convert_ppt_password(self):
        filename = 'ppt_password.ppt'
        expected = 'ppt_password.ppt.txt'
        self.run_test(filename, expected, result=(0, 1), check_magic=False)
        
    def test_convert_pptm(self):
        filename = 'pptm.pptm'
        expected = 'pptm.pptx'
        self.run_test(filename, expected)
        
    def test_convert_odp(self):
        filename = 'odp.odp'
        expected = 'odp.pptx'
        self.run_test(filename, expected)

    def test_convert_msg_clean(self):
        filename = 'clean_attachment.msg'
        expected = 'clean_attachment.pdf'
        self.copy_file(filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + 'clean_attachment.test_attachment_clean.docx'))

    def test_convert_msg_malicious(self):
        filename = 'malicious_attachment.msg'
        expected = 'malicious_attachment.pdf'
        self.copy_file(filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + 'malicious_attachment.doc.docx'))
        
    def test_convert_zip(self):
        filename = 'doc.zip'
        expected = 'doc.zip'
        self.copy_file(filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + 'doc'))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        
        zip = zipfile.ZipFile(target_dir + os.path.sep + expected)
        self.assertFalse('doc.doc' in zip.namelist())
        self.assertTrue('doc.docx' in zip.namelist())
        
    def test_convert_zip_password(self):
        filename = 'doc_password.zip'
        expected = 'doc_password.zip.txt'
        self.copy_file(filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((0, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
    
    # TODO: create a test that checks each file extension in malicious_file_filter
    def test_convert_bat(self):
        filename = 'bat.bat'
        expected = 'bat.bat.txt'
        self.run_test(filename, expected, check_magic=False)
        
    def test_convert_exe(self):
        filename = 'exe.exe'
        expected = 'exe.exe.txt'
        self.run_test(filename, expected, check_magic=False)
    
    

if __name__ == '__main__':
    unittest.main()