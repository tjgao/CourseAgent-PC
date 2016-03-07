import win32api, win32gui, win32con, time
import win32com.client, logging

class win:
    def __init__(self, logger = None):
        self.hwnd = None
        if not logger:
            self.logger = logging.getLogger(__name__)
        else:
            self.logger = logger

    def valid(self):
        if not self.hwnd: return False
        if win32gui.IsWindow(self.hwnd): return True
        self.hwnd = None
        return False

    def _find(self, hwnd, text):
        t = win32gui.GetWindowText(hwnd)
        if text == t.encode('utf-8'):
            self.hwnd = hwnd
            return False
        return True 
    
    def _wildfind(self, hwnd, text):
        t = win32gui.GetWindowText(hwnd)
        if text in t:
            self.hwnd = hwnd
            return False 
        return True

    def find(self, text):
        try: 
            win32gui.EnumWindows(self._find, text)
        except Exception as e:
            if 0 != win32api.GetLastError():
                self.logger.exception(e)

    def wildfind(self, text):
        try:
            win32gui.EnumWindows(self._wildfind, text)
        except Exception as e:
            if 0 != win32api.GetLastError():
                self.logger.exception(e)

    def center(self, zorder=win32con.HWND_TOP):
        if not self.hwnd: return
        try:
            left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
            w, h = right - left, bottom - top
            screenw = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screenh = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            win32gui.SetWindowPos(self.hwnd, zorder, int((screenw - w)/2), int((screenh - h)/2), w, h, win32con.SWP_NOSIZE )
        except Exception as e:
            self.logger.exception(e)

    def activate(self):
        if self.hwnd:
            win32gui.SetActiveWindow(self.hwnd)

    def setfocus(self):
        if self.hwnd:
            win32gui.SetFocus(self.hwnd)

    def pptforeground(self):
        temp = win32gui.GetWindowText (win32gui.GetForegroundWindow())
        if 'PowerPoint' in temp: return True
        return False

    def foreground(self):
        trytimes = 3
        if self.hwnd:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(self.hwnd)
            while not self.pptforeground() and trytimes > 0:
                time.sleep(0.2)
                win32gui.SetForegroundWindow(self.hwnd)

    def maximize(self):
        if self.hwnd:
            win32gui.ShowWindow(self.hwnd, win32con.SW_MAXIMIZE)

    def minimize(self):
        if self.hwnd:
            win32gui.ShowWindow(self.hwnd, win32con.SW_MINIMIZE)

    def exit(self):
        if self.hwnd:
            win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0,0)

    def kill(self):
        if not self.hwnd: return
        try:
            t, p = win32process.GetWindowThreadProcessId(self.hwnd)
            h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, p)
            if h:
                win32api.TerminateProcess(h, 0)
                win32api.CloseHandle(h)
                self.hwnd = None
        except Exception as e:
            self.logger.exception(e)


if __name__ == '__main__':
    import time
    w = win()
    w.wildfind('_._ 图片查看 _._')
    if w.valid():
        w.activate()
        w.foreground()
        w.center()
        #w.maximize()
        #time.sleep(3)
        #w.minimize()
        time.sleep(3)
        w.exit() 
