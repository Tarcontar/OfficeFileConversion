import os
import unittest
import shutil
from os.path import dirname as up
from test_utils import TestCase


class FileConversionTests(TestCase):
    def test_convert_doc(self):
        source = up(self.relpath()) + '\\source\\doc.doc'
        target = self.outpath() + 'doc.doc'
        print(source)
        print(target)
        shutil.copyfile(source, target)
        
        os.system(f'{up(self.relpath())}\\convert_files.py C:\\ {self.outpath()}')
        self.assertTrue(os.path.exists(source + 'x'))
        

if __name__ == '__main__':
    unittest.main()