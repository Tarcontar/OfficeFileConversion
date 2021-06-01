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
        if not os.path.exists(cls.__filepath):
            os.makedirs(cls.__filepath)
        return os.path.join(cls.__filepath, path)
    
    def setUp(self):
        namespace = self.id().split('.')
        testname = namespace[-2] + '.' + namespace[-1]
        print(testname)

        self.__filepath += testname + os.path.sep
        print(self.__filepath)
        if not os.path.exists(self.__filepath):
            os.makedirs(self.__filepath)
            
    def tearDown(self):
        if os.path.exists(self.__filepath):
            shutil.rmtree(self.__filepath)