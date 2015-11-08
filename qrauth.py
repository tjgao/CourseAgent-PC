import qrcode, time, requests, socket, logging, hashlib
from ImgViewer import imgViewer
from threading import Thread

class qrAuth(imgViewer):
    def __init__(self, cfg, logger):
        super().__init__()
        self.cfg= cfg
        self.logger= logger

    def reqAuth(self, code):
        try:
            ip = socket.gethostbyname(socket.gethostname()) 
            port = self.cfg.get('port',9503)
            server = self.cfg.get('server')
            url = server + '/api/qr/create/' + ip + '/' + str(port) + '/' + code 
            r = requests.post(url)
            j = r.json()
            if j.get('code') == 0 :
                return True
            return False
        except:
            self.logger.info('Exception occurs when requesting QR auth')
            return False

    def waitConfirm(self, cycle, code):
        server = self.cfg.get('server')
        url = server + '/api/qr/wait/' + code
        self.logger.info('Wait for confirmation.')
        while True:
            r = requests.get(url)
            try:
                j = r.json()
                if j.get('code') == 0:
                    self.cfg['token'] = j.get('token')
                    m = hashlib.md5()
                    m.update(self.cfg['token'].encode('utf-8'))
                    self.cfg['token2'] = m.hexdigest()
                    self.cfg['authtime'] = j.get('time')
                    self.cfg['sid'] = j.get('sid')
                    self.cfg['uid'] = j.get('uid')
                    self.cfg['authid'] = j.get('authid')
                    self.running = False
                    self.logger.info('Confirmation received!')
                    break
            except Exception as e:
                self.logger.info('Exception occurs when waiting for confirmation')
                self.logger.info(e)
            time.sleep(cycle)

    def show(self, code, title=''):
        if not self.reqAuth(code):
            self.logger.info('Fail to create QR Auth ')
            return
        self.logger.info('Successfully create QR Auth!')
        qr = qrcode.QRCode(version=6)
        qr.add_data(code)
        qr.make()
        img = qr.make_image()
        secs = self.cfg.get('auth_refreshrate',2)
        t = Thread(target=self.waitConfirm, args=(secs,code,))
        t.daemon = True
        t.start()
        self.open(img,'请扫描二维码登入')


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2: sys.exit() 
    q = qrAuth(None)
    q.show(sys.argv[1])
     
