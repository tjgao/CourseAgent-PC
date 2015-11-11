import logging, os, json, atexit, shutil
import cherrypy, multiprocessing, time, hashlib, sys, enum
import ImgViewer, win32gui, win32con, logging, subprocess
from uuid import getnode
from PIL import Image
from queue import Queue
from MSPPT import MSPPT

DETACHED_PROCESS = 8

logger = logging.getLogger(__name__)


viewer_title = '_._ 图片查看 _._'
attend_title = '_._ 课堂签到 _._'

class QTYPE(enum.Enum):
    TIMEOUT = 0
    CHATMSG = 1
    SHOWUP = 2
    SHOWOFF = 3
    RAISEHAND = 4
    GAMEOVER = 1000

class qulet:
    def __init__(self, qtype, data = None):
        self.qtype = qtype
        self.data = data
        self.code = 0
    def __str__(self):
        return json.dumps(self.__dict__)


class courseAgent:
    def __init__(self, gconfig, queue, configurer):
        self.gconfig = gconfig
        self.longpoll = self.gconfig.get('longpoll',10)
        self.queue = queue
        self.configurer = configurer
        self.logger = logging.getLogger(__name__)

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


    def _closewindow(self, title):
        try:
            w = win32gui.FindWindow(title, None)
            if w: win32gui.PostMessage(w, win32con.WM_CLOSE, 0, 0)
            return True
        except:
            return False

    def notifyChange(self, s, fn):
        d = dict()
        d['filename'] = fn
        item = qulet(QTYPE.SHOWUP, d)
        for i, session in s.cache.items():
            session[0].get('queue').put_nowait(item)


    @cherrypy.expose
    def index(self):
        return 'I am alive, ready to go.'


    @cherrypy.expose
    def gameover(self):
        item = qulet(QTYPE.GAMEOVER)
        for i, session in cherrypy.session.cache.items():
            session[0].get('queue').put_nowait(item)
        time.sleep(1)
        cherrypy.engine.exit()

    @cherrypy.expose
    def register(self, token, uid, uname, nickname):
        tk = self.gconfig.get('token')
        if tk != token:
            tk = self.gconfig.get('token2')
            if tk != token:
                return self.jsonify(code=1, msg='No permission')
            else:
                cherrypy.session['admin'] = False
        else:
            cherrypy.session['admin'] = True
        cherrypy.session['uid'] = uid
        cherrypy.session['uname'] = uname
        cherrypy.session['nickname'] = nickname
        cherrypy.session['token'] = token
        cherrypy.session['queue'] = Queue()
        return self.jsonify(code=0)

    @cherrypy.expose
    def chat(self, msg):
        if not cherrypy.session.get('uid'): return self.jsonify(code=1, msg='No Permission')
        d = dict()
        d['uname'] = cherrypy.session.get('uname')
        d['nickname'] = cherrypy.session.get('nickname')
        d['uid'] = cherrypy.session.get('uid')
        d['msg'] = msg
        item = qulet(QTYPE.CHATMSG, d)
        for i, session in cherrypy.session.cache.items():
            session[0].get('queue').put_nowait(item)
        return self.jsonify(code=0)

    @cherrypy.expose
    def watch(self):
        if not cherrypy.session.get('uid') or not cherrypy.session.get('queue'): 
            return self.jsonify(code=1, msg='No Permission')
        q = cherrypy.session.get('queue')
        item = q.get(True, self.longpoll)
        return str(item)
        
        

    @cherrypy.expose
    def currentslide(self):
        if not cherrypy.session.get('uid'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            return self.jsonify(idx=self.ppt.curslide(), code=0)
        return self.jsonify(code=1, msg='Powerpoint is not running')

    @cherrypy.expose
    def raisehand(self, filename):
        if not cherrypy.session.get('uid'): return self.jsonify(code=1, msg='No Permission')
        d = dict()
        d['uname'] = cherrypy.session.get('uname')
        d['nickname'] = cherrypy.session.get('nickname')
        d['uid'] = cherrypy.session.get('uid')
        d['filename'] = filename
        item = qulet(QTYPE.RAISEHAND, d)
        for i, session in cherrypy.session.cache.items():
            session[0].get('queue').put_nowait(item)
        return self.jsonify(code=0)

    @cherrypy.expose
    def closeviewer(self):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        self._closewindow(viewer_title)
        self._closewindow(attend_title)
        return self.jsonify(code=0)
        
    @cherrypy.expose
    def showattend(self, attendstr):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        self._closewindow(attend_title) 
        subprocess.Popen(['python','ImgViewer.py','-q',attendstr,viewer_title],creationflags=DETACHED_PROCESS,close_fds=True) 
        return self.jsonify(code=0)

    @cherrypy.expose
    def upload(self,filename):
        if not cherrypy.session.get('uid'):
            return self.jsonify(code=1, msg='No permission')

        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        if not os.path.exists(path):
            os.mkdir(path)
        m = hashlib.md5()
        ext = ''
        if '.' in filename:
            ext = '.' + filename.split('.')[-1]
        m.update((filename + str(time.time())).encode('utf-8'))
        dest = path + os.sep + m.hexdigest() + ext.lower()
        with open(dest, 'wb') as f:
            shutil.copyfileobj(cherrypy.request.body, f)
        if ext.lower() in ['.jpg','.png','.gif','.bmp']:
            thumb = self.thumb(dest)
            return self.jsonify(thumb=os.path.basename(thumb), filename=os.path.basename(dest), code=0)
        else:
            return self.jsonify(filename = os.path.basename(dest), code=0)

    @cherrypy.expose
    def openppt(self, filename):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt:
            if self.ppt.opened():
                self.ppt.close()
        else:
            self.ppt = MSPPT()
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        self.ppt.open(path + os.sep + filename)
        if not os.path.exists(path):
            os.mkdir(path)
        self.ppt.makepics( path ) 
        # slide's name is 1.png,2.png,....
        self.notifyChange(cherrypy.session, 'files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0) 

    @cherrypy.expose
    def ppt_next(self):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.next()
            self.notifyChange(cherrypy.session, 'files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_prev(self):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.prev()
            self.notifyChange(cherrypy.session, 'files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_pages(self):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            return self.jsonify(code=0, data=self.ppt.pages())
        return self.jsonify(code=1, msg='PPT is not running')

    @cherrypy.expose
    def ppt_goto(self, page):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.goto(page)
            self.notifyChange(cherrypy.session, 'files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)


    @cherrypy.expose
    def ppt_first(self):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.first()
            self.notifyChange(cherrypy.session, 'files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_last(self):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.last()
            self.notifyChange(cherrypy.session, 'files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)
        
    @cherrypy.expose
    def show(self, filename):
        if not cherrypy.session.get('admin'):
            return self.jsonify(code=1, msg='No permission')

        # start a image viewer to open the picture 
        self._closewindow(viewer_title)
        subprocess.Popen(['python','ImgViewer.py','-f',filename,viewer_title],creationflags=DETACHED_PROCESS,close_fds=True) 
        return self.jsonify(code=0)

if __name__ == '__main__':
    pass
