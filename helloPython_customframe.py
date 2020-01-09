import os
from clr import *
from MapInfo.Types import PythonAddinCustomFrame, DimensionedValue, PaperUnits, InactiveLayoutContent, UnitConversion
from mi_addin_layout_customframe_util import LayoutCustomFrameUtil
from mi_common_util import CommonUtil
from System import Double

class CustomFrameControl(PythonAddinCustomFrame):
    __assembly__  = os.path.basename(__file__)
    def __init__(self):
        self._active = False
        self._addin_mbx = ""
        self._brush_clause = ""
        self._can_activate = False
        self._can_remove_content = False
        self._can_resize = False
        self._frame_height = DimensionedValue(0, PaperUnits.Inches)
        self._frame_name = self. __assembly__
        self._frame_width = DimensionedValue(0, PaperUnits.Inches)
        self._layout_page_scale = float(0)
        self._name = ""
        self._pen_border_clause = ""
        self._preseve_aspect = False
        self._xcoord = DimensionedValue(0, PaperUnits.Inches)
        self._ycoord = DimensionedValue(0, PaperUnits.Inches)
        self._zindex = 0
        self._content = None
        self._wrapper = None
        self._disposed = False

    @property
    def WrapperControl(self):
        return self._wrapper
    @WrapperControl.setter
    def WrapperControl(self, value):
        self._wrapper = value

    @property
    def ControlContent(self):
        return self._content
    @ControlContent.setter
    def ControlContent(self, value):
        self._content = value

    @clrmethod(None, [bool])
    def Dispose(self, disposing):
        if not self._disposed:
            try:
                if self.WrapperControl and disposing:
                    self.WrapperControl.Dispose()
                    self.WrapperControl = None
            finally:
                self._disposed = True
                super().Dispose(disposing)

    @clrmethod(InactiveLayoutContent, [Double])
    def PrintContent(self, printScale):
        try:
            if self.ControlContent:
                tempScale = self.LayoutPageScale
                self.LayoutPageScale = printScale
                printContent = InactiveLayoutContent(self.ControlContent)
                self.LayoutPageScale = tempScale
                return printContent
        except Exception as e:
            CommonUtil.sprint("Failed to Print: {}".format(e))
        return None

    @clrmethod(None, [str])
    def Serialize(self, fileName):
        LayoutCustomFrameUtil.serialize_frame(self, fileName)
    
    @clrproperty(int)
    def ZIndex(self):
        return self._zindex
    @ZIndex.setter
    def ZIndex(self, value):
        self._zindex = value

    @clrproperty(DimensionedValue)
    def YCoord(self):
        return self._ycoord
    @YCoord.setter
    def YCoord(self, value):
        self._ycoord = UnitConversion.ConvertUnits(value, PaperUnits.Pixels)

    @clrproperty(DimensionedValue)
    def XCoord(self):
        return self._xcoord
    @XCoord.setter
    def XCoord(self, value):
        self._xcoord = UnitConversion.ConvertUnits(value, PaperUnits.Pixels)

    @clrproperty(bool)
    def PreserveAspect(self):
        return self._preseve_aspect
    @PreserveAspect.setter
    def PreserveAspect(self, value):
        self._preseve_aspect = value

    @clrproperty(str)
    def PenBorderClause(self):
        return self._pen_border_clause
    @PenBorderClause.setter
    def PenBorderClause(self, value):
        self._pen_border_clause = value

    @clrproperty(str)
    def Name(self):
        return self._name
    @Name.setter
    def Name(self, value):
        self._name = value

    @clrproperty(Double)
    def LayoutPageScale(self):
        return self._layout_page_scale
    @LayoutPageScale.setter
    def LayoutPageScale(self, value):
        self._layout_page_scale = value

    @clrproperty(DimensionedValue)
    def FrameWidth(self):
        return self._frame_width
    @FrameWidth.setter
    def FrameWidth(self, value):
        self._frame_width = UnitConversion.ConvertUnits(value, PaperUnits.Pixels)

    @clrproperty(str)
    def FrameName(self):
        return self._frame_name
    @FrameName.setter
    def FrameName(self, value):
        self._frame_name = value

    @clrproperty(DimensionedValue)
    def FrameHeight(self):
        return self._frame_height
    @FrameHeight.setter
    def FrameHeight(self, value):
        self._frame_height = UnitConversion.ConvertUnits(value, PaperUnits.Pixels)

    @clrproperty(bool)
    def Active(self):
        return self._active
    @Active.setter
    def Active(self, value):
        self._active = value
    
    @clrproperty(str)
    def AddinMBXFileName(self):
        return self._addin_mbx
    @AddinMBXFileName.setter
    def AddinMBXFileName(self, value):
        self._addin_mbx = value
    
    @clrproperty(str)
    def BrushClause(self):
        return self._brush_clause
    @BrushClause.setter
    def BrushClause(self, value):
        self._brush_clause = value

    @clrproperty(bool)
    def CanActivate(self):
        return self._can_activate
    @CanActivate.setter
    def CanActivate(self, value):
        self._can_activate = value

    @clrproperty(bool)
    def CanRemoveContent(self):
        return self._can_remove_content
    @CanRemoveContent.setter
    def CanRemoveContent(self, value):
        self._can_remove_content = value

    @clrproperty(bool)
    def CanResize(self):
        return self._can_resize
    @CanResize.setter
    def CanResize(self, value):
        self._can_resize = value
    