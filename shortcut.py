import win32com.client
import os
from win32com.shell import shell, shellcon


desktop = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
path = os.path.dirname(os.path.realpath(__file__)) 
#print( 'path:' + path ) 
ws = win32com.client.Dispatch('wscript.shell')
scut = ws.CreateShortcut(desktop + os.sep + 'SACS.lnk')
scut.TargetPath = '"pythonw.exe"'
scut.Arguments = path + os.sep + 'main.py'
scut.WorkingDirectory = path
scut.Save()


