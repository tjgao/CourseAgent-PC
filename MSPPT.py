# A simple class for powerpoint manuipulation
# Author: Tiejun Gao
import win32com.client, time
from atlogger import logger


class MSPPT:
    def __init__(self, app = None):
        self.app = app
        self.ppt = None
        self.DEBUG = True

    def open(self, target):
       #Before we actually open ppt, we try to minimize all other windows
        try:
            sh = win32com.client.Dispatch('Shell.Application')
            sh.minimizeAll();
            if self.app is None:
                self.app = MSPPT.getRunning()
                if self.app is None:
                    self.app = win32com.client.Dispatch('PowerPoint.Application')
            self.app.Visible = True
            self.app.WindowState =  win32com.client.constants.ppWindowMaximized
            self.app.Activate()
            self.ppt = self.app.Presentations.Open(FileName = target, ReadOnly = True)
            self.ppt.SlideShowSettings.Run()
        except Exception as e:
            logger.exception(e)

    def getRunning():
        try:
            r = win32com.client.GetActiveObject('PowerPoint.Application')
            return r
        except Exception as e:
            logger.exception(e)
            return None

    def next(self):
        if self.ppt is None: return
        try:
            self.ppt.SlideShowWindow.View.Next()
        except Exception as e:
            logger.exception(e)

    def prev(self):
        if self.ppt is None: return
        try:
            self.ppt.SlideShowWindow.View.Previous()
        except Exception as e:
            logger.exception(e)

    def pages(self):
        if self.ppt is None: return 0
        return len(self.ppt.Slides)

    def goto(self, idx):
        if self.ppt is None: return
        if idx > len(self.ppt.Slides) or idx < 1: return
        try:
            self.ppt.SlideShowWindow.View.GotoSlide(idx)
        except Exception as e:
            logger.exception(e)

    def last(self):
        if self.ppt is None: return
        try:
            self.ppt.SlideShowWindow.View.Last()
        except Exception as e:
            logger.exception(e)

    def first(self):
        if self.ppt is None: return
        try:
            self.ppt.SlideShowWindow.View.First()
        except Exception as e:
            logger.exception(e)
            


    def close(self):
        try:
            if self.ppt is not None:
                self.ppt.SlideShowWindow.View.Exit()
            if self.app is not None:
                self.app.Quit()
            self.app, self.ppt = None, None
        except Exception as e:
            logger.exception(e)


import time
if __name__ == '__main__':
    r = MSPPT.getRunning()
    ppt = None
    if r is None: 
        print('Powerpoint is not running')
        ppt = MSPPT()
    else:
        ppt = MSPPT(r)
    ppt.open('d:\\code\\test.pptx')
    r = MSPPT.getRunning()
    if r is not None: print('Powerpoint running')
    print('ppt slides:' + str(ppt.pages()))
    ppt.goto(3)
    #time.sleep(10)
    #ppt.close()
