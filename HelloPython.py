# Sample addin as example of creating MapInfo Pro addin in Python.
# This sample does the following things.
# 1.) Creates a ribbon tab, ribbon button and a tool button.
# 2.) Shows how to attach a command to button firing a python function.
# 3.) Shows how to create and add a custom frame into layout window.
# 4.) Shows how to call a mapbasic subroutine or function from python code.
#
# How To Run:
#   1.) Compile and link the helloPython.mb file to helloPython.mbx.
#   2.) Run the helloPython.mbx.
#
# Note : Sample is only for MapInfo Pro 64 bit v17.0.3 or greater.

import ptvsd
# ptvsd.enable_attach()
# ptvsd.wait_for_attach()

import sys
import os
from os.path import join, dirname
from osgeo import ogr

sys.path.append(join(dirname(__file__), "C:\\Program Files\\MapInfo\\Professional"))
sys.path.append(join(dirname(__file__), "py_addin_util"))

pathMsg = ""
for ele in sys.path:
    pathMsg += ele
    
print(pathMsg)

import clr # How do I import this MapInfo python library?
clr.AddReference(r"WindowsBase")
from mi_common_util import CommonUtil
from mi_addin_util import AddinUtil
from mi_addin_layout_customframe_util import LayoutCustomFrameUtil
from mi_addin_resourceManager import StringResourceManager
from helloPython_customframe import CustomFrameControl
from MapInfo.Types import IMapInfoPro, ControlType, PaperUnits, DimensionedValue, NotificationObject, NotificationType
from System.Windows import Size, Point
from System import Double


# this class is needed with same name in order to load the python addin and can be copied 
# as it is when creating another addin.
class MyAddin():
    def __init__(self, imapinfopro, thisApplication):
        try:
            self._view_file = join(dirname(__file__), "helloPython_customframe_view.xaml")
            self._pro = imapinfopro
            self._thisApplication = thisApplication
            self._stringResourceManager = StringResourceManager(thisApplication)

            CommonUtil.sprint("loading...")
            self._newTab = self._pro.Ribbon.Tabs.Add(self.get_resource_string(4))

            if thisApplication:
                # this is how we can call a mbx application subroutine from python.
                thisApplication.CallMapBasicSubroutine("FromPython", None)
                CommonUtil.do("""
print "Hello from (MapBasic code inside) Python Addin!"
                """)
            
            # don't use end mapinfo mapbasic statement in python instead use CommonUtil.end_application().
            # CommonUtil.end_application()

        except Exception as e:
            CommonUtil.sprint("Failed to load: {}".format(e))

    def get_resource_string(self, id) -> str:
        if self._stringResourceManager:
            return self._stringResourceManager.GetResourceString(id)

    def unload(self):
        CommonUtil.sprint("unloading...")
        
        if self._newTab:
            self._pro.Ribbon.Tabs.Remove(self._newTab)
        if self._stringResourceManager:
            del self._stringResourceManager
        self._stringResourceManager = None
        self._newTab = None
        self._thisApplication = None
        self._pro = None

class main():
    def __init__(self, imapinfopro):
        self._imapinfopro = imapinfopro

    def load(self):
        try:
            # uncomment these lines to debug the python script using VSCODE
            # Install ptvsd package into your environment using "pip install ptvsd"
            # Debug in VSCODE with Python: Attach configuration

            # here initialize the addin class
            if self._imapinfopro:
                # obtain the handle to current application if needed
                # thisApplication = self._imapinfopro.GetMapBasicApplication(os.path.splitext(__file__)[0] + ".mbx")
                # self._addin = MyAddin(self._imapinfopro, thisApplication)
                self._addin = None
                CommonUtil.sprint("Hello from Python Addin!")
                CommonUtil.do("""
print "Hello from (MapBasic code inside) Python Addin!"
                """)
                CommonUtil.sprint("sys.path => " + pathMsg)

        except Exception as e:
            CommonUtil.sprint("Failed to load: {}".format(e))
    
    def reconstruct(self, fileName, xPosition, yPosition, width, height, penMapBasicCommand, brushMapBasicCommand, zIndex, name):
        try:
            CommonUtil.sprint("Reconstruct addin")
        except Exception as e:
            CommonUtil.sprint("Failed to reconstruct: {}".format(e))  

    def unload(self):
        try:
            CommonUtil.sprint("Unload addin")
            self._addin = None
        except Exception as e:
            CommonUtil.sprint("Failed to unload: {}".format(e))
    
    def __del__(self):
        self._imapinfopro = None
        pass
