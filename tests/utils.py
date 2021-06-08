import os
import tempfile
import shutil
import inspect
import unittest

class TestCase(unittest.TestCase):
    __filepath = os.path.join(tempfile.gettempdir(), 'tests_tmp')

    @classmethod
    def relpath(cls, path=None):
        result = os.path.dirname(inspect.getfile(cls))
        if path is not None:
            result = os.path.join(result, path)
        return result
    
    @classmethod
    def outpath(cls, path=''):
        new_path = os.path.join(cls.__filepath, path)
        if not os.path.exists(new_path):
            os.makedirs(new_path)
        return new_path
    
    def setUp(self):
        namespace = self.id().split('.')
        testname = namespace[-2] + '.' + namespace[-1]
        print(testname)
        
        if os.path.exists(self.__filepath):
            shutil.rmtree(self.__filepath)
            
        os.makedirs(self.__filepath)
            
    def tearDown(self):
        if os.path.exists(self.__filepath):
            shutil.rmtree(self.__filepath)