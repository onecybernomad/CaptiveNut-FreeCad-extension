import os
import FreeCAD
import FreeCADGui

# Register icon path at import time
icon_dir = os.path.join(FreeCAD.getUserAppDataDir(), "Mod", "CaptiveNut", "Resources", "icons")
if os.path.isdir(icon_dir):
    FreeCADGui.addIconPath(icon_dir)

class CaptiveNutWorkbench(FreeCADGui.Workbench):
    MenuText = "CaptiveNut"
    Icon = "CaptiveNut.svg"
    def Initialize(self):
        import CaptiveNut
        self.appendToolbar("CaptiveNut", ["CaptiveNut_Command"])
    def GetClassName(self):
        return "Gui::PythonWorkbench"

FreeCADGui.addWorkbench(CaptiveNutWorkbench())