import logging, os, json, atexit, shutil
import cherrypy, multiprocessing, time, hashlib, sys
from uuid import getnode
from PIL import Image

logger = logging.getLogger(__name__)


class courseAgent:
    def __init__(self, gconfig):
        self.cfg = gconfig

    def thumb(self, src):
        try:
            sz = gconfig.get('thumb_size','50x50')
            a,b = sz.split('x')
            width, height = int(a), int(b) 
        except:
            logger.info('Thumb size format wrong, use default value 50x50')
            width, height = 50, 50
        dirname = os.path.dirname(src)
        basename = os.path.basename(src)
        dst = dirname + os.sep + 'tb_' + basename
        Image.open(src).thumbnail((width,height)).save(dst)        
        return dst

    def jsonify(self, **args):
        return json.dumps(args)

    @cherrypy.expose
    def index(self):
        pass

    @cherrypy.expose
    def takeover(self, token):
        pass
        

    @cherrypy.expose
    def upload(self, token, filename):
        tk = self.gconfig.get('token')
        if tk != token:
            tk = self.gconfig.get('token2')
            if tk != token:
                return self.jsonify(code=1, msg='No permission')

        path =  os.path.dirname(os.path.realpath(__file__)) + os.sep + 'upload'
        if not os.path.exists(path):
            os.mkdir(path)
        m = hashlib.md5()
        ext = ''
        if '.' in filename:
            ext = '.' + filename.split('.')[-1]
        m.update((filename + str(time.time())).encode('utf-8'))
        dest = path + os.sep + m.hexdigest() + ext
        with open(dest, 'wb') as f:
            shutil.copyfileobj(cherrypy.request.body, f)
        thumb = self.thumb(dest)
        return self.jsonify(thumb=os.path.basename(thumb), filename=os.path.basename(dest), code=0)

    @cherrypy.expose
    def project(self, token, filename):
        tk = self.gconfig.get('token')
        if tk != token:
            return self.jsonify(code=1, msg='No permission')

        # start a image viewer to open that picture 
        return self.jsonify(code=0)

if __name__ == '__main__':
    pass
