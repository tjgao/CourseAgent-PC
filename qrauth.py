import qrcode, time
from ImgViewer import imgViewer
from threading import Thread
class qrAuth(imgViewer):
    def __init__(self, st):
        super().__init__()
        self.auth = None
        self.setting = st 

    def getAuth(self):
        return self.auth

    def waitConfirm(self, cycle):
        while True:
            time.sleep(cycle)
            #self.auth = '12345'
            #self.running = False

    def show(self, code, title=''):
        qr = qrcode.QRCode(version=6)
        qr.add_data(code)
        qr.make()
        img = qr.make_image()
        t = Thread(target=self.waitConfirm, args=(3,))
        t.daemon = True
        t.start()
        self.open(img,'')

import sys

if __name__ == '__main__':
    if len(sys.argv) < 2: sys.exit() 
    q = qrAuth(None)
    q.show(sys.argv[1])
     
