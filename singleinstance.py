import win32event, win32api, winerror
from uuid import getnode

class singleInstance:
    def __init__(self):
        self.muName = '_CA_MUTEX_' + str(getnode()) 
        self.mutex = win32event.CreateMutex(None, False, self.muName)
        self.lasterror = win32api.GetLastError()

    def alreadyRunning(self):
        return self.lasterror == winerror.ERROR_ALREADY_EXISTS

