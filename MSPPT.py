# A simple class for powerpoint manuipulation
import win32com.client, time, os, logging
import pythoncom, winreg


class MSPPT:
    def __init__(self, logger=None, app = None):
        self.app = app
        self.ppt = None
        if logger:
            self.logger = logger
        self.logger = logging.getLogger(__name__)
        self.pages = 0
        # to support multi-threaded model, com objects have to be marshalled
        self.marshal_app = None
        self.marhsal_ppt = None

    def opened(self):
        return self.marshal_app is not None 


    def _getpptexe(self):
        def _regkey_value(path, name="", start_key = None):
            if isinstance(path, str):
                path = path.split("\\")
            if start_key is None:
                start_key = getattr(winreg, path[0])
                return _regkey_value(path[1:], name, start_key)
            else:
                subkey = path.pop(0)
            with winreg.OpenKey(start_key, subkey) as handle:
                assert handle
                if path:
                    return _regkey_value(path, name, handle)
                else:
                    desc, i = None, 0
                    while not desc or desc[0] != name:
                        desc = winreg.EnumValue(handle, i)
                        i += 1
                    return desc[1]
        val = _regkey_value(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe", "")
        return os.path.basename(val)



    def kill(self, image):
        os.system('taskkill /F /im ' + image)

    def open(self, target):
       #Before we actually open ppt, we try to minimize all other windows
        try:
            pythoncom.CoInitialize()
            r = MSPPT.getRunning()
            if r is not None:
                r.Quit()
            #sh = win32com.client.Dispatch('Shell.Application')
            #sh.minimizeAll();
            self.app = win32com.client.Dispatch('PowerPoint.Application')
            self.app.Visible = True
            #self.app.WindowState =  win32com.client.constants.ppWindowMaximized
            self.app.WindowState = 3
            self.app.Activate()
            self.ppt = self.app.Presentations.Open(FileName = target, ReadOnly = True)
            self.pages = len(self.ppt.Slides)
            self.ppt.SlideShowSettings.Run()
            # marshal these com objects
            self.marshal_app = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch, self.app
                    )
            #self.marshal_ppt = pythoncom.CoMarshalInterThreadInterfaceInStream(
            #        pythoncom.IID_IDispatch, self.ppt
            #        )
            #self.marshal_ppt = pythoncom.CreateStreamOnHGlobal()
            pythoncom.CoMarshalInterface(self.marshal_app,
                                         pythoncom.IID_IDispatch,
                                         self.app._oleobj_,
                                         pythoncom.MSHCTX_LOCAL,
                                         pythoncom.MSHLFLAGS_TABLESTRONG)

        except Exception as e:
            print(e)
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize()

    def getRunning():
        try:
            r = win32com.client.GetActiveObject('PowerPoint.Application')
            return r
        except Exception as e:
            return None

    def getpages(self): return self.pages

    def next(self):
        if self.marshal_app is None: return
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            if app.Presentations(1).SlideShowWindow.View.Slide.SlideIndex >= self.pages: return
            app.Presentations(1).SlideShowWindow.View.Next()
        except Exception as e:
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize();

    def prev(self):
        if self.marshal_app is None: return
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            if app.Presentations(1).SlideShowWindow.View.Slide.SlideIndex <= 1: return
            app.Presentations(1).SlideShowWindow.View.Previous()
        except Exception as e:
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize();

    def goto(self, idx):
        if self.marshal_app is None: return
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            app.Presentations(1).SlideShowWindow.View.GotoSlide(idx)
        except Exception as e:
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize();

    def last(self):
        if self.marshal_app is None: return
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            app.Presentations(1).SlideShowWindow.View.Last()
        except Exception as e:
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize();

    def first(self):
        if self.marshal_app is None: return
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            app.Presentations(1).SlideShowWindow.View.First()
        except Exception as e:
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize();

    # makepics usually will only be called when opening a ppt, in the same thread
    def makepics(self, path):
        if self.marshal_app is None: return 0
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            ppt = app.Presentations(1)
            for i, s in enumerate(ppt.Slides):
                s.Export( path + os.sep + str(i+1) + '.png', 'png')
        except Exception as e:
            self.logger.exception(e)
        finally:
            pythoncom.CoUninitialize();

    def close(self):
        try:
            pythoncom.CoInitialize()
            r = MSPPT.getRunning()
            if r: r.Quit()
            if self.marshal_app:
                self.marshal_app.Seek(0,0)
                app = win32com.client.Dispatch(
                        pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
                app.Quit()
        except Exception as e:
            self.logger.exception(e)
        finally:
            self.marshal_app = None
            pythoncom.CoUninitialize();
            self.kill(self._getpptexe())


    def curslide(self):
        if self.marshal_app is None: return 0
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            return app.Presentations(1).SlideShowWindow.View.Slide.SlideIndex
        except Exception as e:
            self.logger.exception(e)
            return 0
        finally:
            pythoncom.CoUninitialize();

    def gethwnd(self):
        if self.marshal_app is None: return None
        try:
            pythoncom.CoInitialize()
            self.marshal_app.Seek(0,0)
            app = win32com.client.Dispatch(
                    pythoncom.CoUnmarshalInterface(self.marshal_app, pythoncom.IID_IDispatch)
                    )
            
        except:
            pass
        finally:
            pythoncom.CoUninitialize()
import time
if __name__ == '__main__':
    r = MSPPT.getRunning()
    ppt = None
    if r is None: 
        print('Powerpoint is not running')
        ppt = MSPPT()
    else:
        r.Quit()
        print('Quit Running PPT')
        ppt = MSPPT()
    ppt.open('d:\\code\\test.pptx')
    r = MSPPT.getRunning()
    if r is not None: print('Powerpoint running')
    #print('ppt slides:' + str(ppt.getpages()))
    print('current slide:' + str(ppt.curslide()))
    if ppt.opened(): print('opened')
    time.sleep(3)
    ppt.next()
    ppt.makepics('d:\\code\\temp')
    time.sleep(5)
    ppt.close()
