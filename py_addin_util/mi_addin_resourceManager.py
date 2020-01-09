import clr
from MapInfo.Types import IMapBasicApplication

class StringResourceManager:
    def __init__(self, this_application_instance):
        self._thisApplication = IMapBasicApplication(this_application_instance)
        if not self._thisApplication:
            raise Exception("Incorrect Application Instance!")

    def GetResourceString(self, id):
        retStr = ""
        try:
            if self._thisApplication:
                retStr = self._thisApplication.EvaluateMapBasicFunction("GetResString", [str(id)])
        except:
            retStr = ""
        return retStr

    def __del__(self):
        self._thisApplication = None
        pass
