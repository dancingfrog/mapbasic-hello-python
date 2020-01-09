import clr
from MapInfo.Types import PythonProUtil
from System import Uri, UriKind

class CommonUtil:
    def __init__(self):
        pass

    @staticmethod
    def sprint(string):
        PythonProUtil.Print(string)

    @staticmethod
    def do(command):
        PythonProUtil.Do(command)
    
    @staticmethod
    def eval(command):
        try:
            return PythonProUtil.Eval(command)
        except:
            pass
    
    @staticmethod
    def end_application():
        PythonProUtil.EndApplication()
            
    @staticmethod
    def path_to_uri(url_string):
        try:
            return Uri(url_string, UriKind.RelativeOrAbsolute)
        except:
            pass
    
    @staticmethod
    def get_mi_directory() -> str:
        try:
            return PythonProUtil.GetMapInfoProDirectoryPath()
        except:
            return ""