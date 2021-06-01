import os
import sys
import unittest
import shutil
from os.path import dirname as up
from test_utils import TestCase

sys.path.insert(1, '..')
from convert_files import process_folder, get_magic, ZIP_FILE_MAGIC


class FileConversionTests(TestCase):
    def test_convert_doc(self):
        filename = 'doc.doc'
        expected = 'doc.docx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_docm(self):
        filename = 'docm.docm'
        expected = 'docm.docx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_dot(self):
        filename = 'dot.dot'
        expected = 'dot.dotx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_docx(self):
        filename = 'docx.docx'
        expected = 'docx.docx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((0, 0), process_folder(issue_dir, target_dir))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_dot_fake_dotx(self):
        filename = 'dot_fake_dotx.dotx'
        expected = 'dot_fake_dotx.dotx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_dotm(self):
        filename = 'dotm.dotm'
        expected = 'dotm.dotx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_odt(self):
        filename = 'odt.odt'
        expected = 'odt.docx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        self.assertTrue(ZIP_FILE_MAGIC in get_magic(target_dir + os.path.sep + expected))
        
    def test_convert_doc_password(self):
        filename = 'doc_password.doc'
        expected = 'doc_password.doc.txt'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((1, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))
        self.assertTrue(os.path.exists(target_dir + os.path.sep + expected))
        
    def test_convert_autosave(self):
        filename = '~$autosave.docx'
        source = up(self.relpath()) + '\\source'
        source_file = source + os.path.sep + filename
        shutil.copyfile(source_file, self.outpath('source') + os.path.sep + filename)
        
        issue_dir = self.outpath('issue')
        target_dir = self.outpath('source')
        self.assertEqual((0, 0), process_folder(issue_dir, target_dir))
        self.assertFalse(os.path.exists(target_dir + os.path.sep + filename))

        

if __name__ == '__main__':
    unittest.main()