import os
import sys
import unittest
import shutil
from os.path import dirname as up
from test_utils import TestCase

sys.path.insert(1, '..')
from convert_files import process_folder


class FileConversionTests(TestCase):
    def test_convert_doc(self):
        source = up(self.relpath()) + '\\source'
        source_file = source + '\\doc.doc'
        target = self.outpath() + 'doc.doc'
        shutil.copyfile(source_file, target)
        
        issue_dir = self.outpath() + 'issue'
        os.makedirs(issue_dir)
        process_folder(issue_dir, self.outpath())
        self.assertTrue(os.path.exists(self.outpath() + 'doc.docx'))
        

if __name__ == '__main__':
    unittest.main()