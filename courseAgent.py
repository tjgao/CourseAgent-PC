import logging, os, json, atexit, shutil
import cherrypy, multiprocessing, time, hashlib, sys, enum, requests
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
        self._user = {}
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

    def notifyChange(self, fn):
        d = dict()
        d['filename'] = fn
        item = qulet(QTYPE.SHOWUP, d)
        for i in self._user:
            self._user[i].get('queue').put_nowait(item)

    def checktoken(self):
        return 1,2
        uid = cherrypy.request.headers.get('id')
        tok = cherrypy.request.headers.get('token')
        if id is None: return None
        if self._user.get(uid) is None: return None
        if tok == gconfig.get('token'): return uid, 2
        elif tok == gconfig.get('token2'): return uid, 1
        return None

    @cherrypy.expose
    def index(self):
        return 'I am alive, ready to go.'

    @cherrypy.expose
    def alive(self):
        return self.jsonify(code=0)


    @cherrypy.expose
    def gameover(self):
        item = qulet(QTYPE.GAMEOVER)
        for i in self._user:
            self._user[i].get('queue').put_nowait(item)
        time.sleep(1)
        cherrypy.engine.exit()

    @cherrypy.expose
    def register(self, uname, nickname, headimg = None):
        ret = self.checktoken()
        if ret is None: return self.jsonify(code=1)
        uid, rank = ret
        self._user[uid] = {} 
        self._user[uid]['uid'] = uid
        self._user[uid]['uname'] = uname
        self._user[uid]['nickname'] = nickname
        self._user[uid]['queue'] = Queue()
        self._user[uid]['rank'] = rank
        return self.jsonify(code=0)

    @cherrypy.expose
    def chat(self, msg):
        ret = self.checktoken()
        if ret is None: return self.jsonify(code=1)
        d = dict()
        uid, rank = ret
        d['uname'] = self._user[uid].get('uname')
        d['nickname'] = self._user[uid].get('nickname')
        d['uid'] = self._user[uid].get('uid')
        d['headimg'] = self._user[uid].get('headimg')
        d['text'] = msg
        item = qulet(QTYPE.CHATMSG, d)
        for i in self._user:
            self._user[i].get('queue').put_nowait(item)
        return self.jsonify(code=0)

    @cherrypy.expose
    def watch(self):
        ret = self.checktoken()
        if ret is None: return self.jsonify(code=1)
        q = self._user[ret.uid].get('queue')
        item = q.get(True, self.longpoll)
        return str(item)
        
        

    @cherrypy.expose
    def currentview(self):
        ret = self.checktoken()
        if ret is None: return self.jsonify(code=1)
        if self.ppt and self.ppt.opened():
            return self.jsonify(idx=self.ppt.curslide(), code=0)
        return self.jsonify(code=1, msg='Powerpoint is not running')

    @cherrypy.expose
    def raisehand(self, filename):
        ret = self.checktoken()
        if ret is None: return self.jsonify(code=1)
        uid, rank = ret
        d = dict()
        d['uname'] = self._user[uid].get('uname')
        d['nickname'] = self._user[uid].get('nickname')
        d['uid'] = self._user[uid].get('uid')
        d['headimg'] = self._user[uid].get('headimg')
        d['filename'] = filename
        d['text'] = '举手申请展示图片。'
        item = qulet(QTYPE.RAISEHAND, d)
        for i in self._user:
            self._user[i].get('queue').put_nowait(item)
        return self.jsonify(code=0)

    @cherrypy.expose
    def closeviewer(self):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        self._closewindow(viewer_title)
        self._closewindow(attend_title)
        return self.jsonify(code=0)
        
    @cherrypy.expose
    def showattend(self, attendstr):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        self._closewindow(attend_title) 
        subprocess.Popen(['python','ImgViewer.py','-q',attendstr,viewer_title],creationflags=DETACHED_PROCESS,close_fds=True) 
        return self.jsonify(code=0)

    @cherrypy.expose
    def permitshow(self, filename):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        if filename.endswith('.pptx') or filename.endswith('.ppt'):
            self._curppt = path + os.sep + filename
        else:
            self._curpic = path + os.sep + filename
        return self.jsonify(code=0)

    @cherrypy.expose
    def upload(self,filename):
        ret = self.checktoken() 
        if ret is None:
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
        if ext.lower() in ['.jpg','.png','.gif','.bmp','.jpeg']:
            thumb = self.thumb(dest)
            return self.jsonify(thumb=os.path.basename(thumb), filename=os.path.basename(dest), code=0)
        elif ext.lower() in ['.ppt','.pptx']:
            self._curppt = dest
            return self.jsonify(filename = os.path.basename(dest), code=0)
        else:
            return self.jsonify(code=1, msg='只能上传图片或者Powerpoint文件！')

    @cherrypy.expose
    def downloadppt(self, url):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        fn = url.split('/')[-1]
        path =  os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        fn = path + os.sep + fn 
        r = requests.get(url, stream=True)
        with open(fn, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024):
                if chunk: f.write(chunk)
        self._curppt = fn
        return self.jsonify(code=0)

    @cherrypy.expose
    def openppt(self, filename = None):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt:
            if self.ppt.opened():
                self.ppt.close()
        else:
            self.ppt = MSPPT()
        if filename is None and self._curppt is None:
            return self.jsonify(code=1, msg='尚未指定Powerpoint文件！')
        elif self._curppt is not None: self._curppt = filename
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        self.ppt.open(path + os.sep + self._curppt)
        if not os.path.exists(path):
            os.mkdir(path)
        self.ppt.makepics( path ) 
        # slide's name is 1.png,2.png,....
        self.notifyChange('files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0) 

    @cherrypy.expose
    def ppt_next(self):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')

        if self.ppt and self.ppt.opened():
            self.ppt.next()
            self.notifyChange('files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_prev(self):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')

        if self.ppt and self.ppt.opened():
            self.ppt.prev()
            self.notifyChange('files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_pages(self):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            return self.jsonify(code=0, data=self.ppt.pages())
        return self.jsonify(code=1, msg='PPT is not running')

    @cherrypy.expose
    def ppt_goto(self, page):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.goto(page)
            self.notifyChange('files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)


    @cherrypy.expose
    def ppt_first(self):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.first()
            self.notifyChange('files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_last(self):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.last()
            self.notifyChange('files' + os.sep + self.ppt.curslide() + '.png')
        return self.jsonify(code=0)
        
    @cherrypy.expose
    def show(self, filename):
        ret = self.checktoken() 
        if ret is None or ret.rank < 2:
            return self.jsonify(code=1, msg='No permission')

        # start a image viewer to open the picture 
        self._closewindow(viewer_title)
        subprocess.Popen(['python','ImgViewer.py','-f',filename,viewer_title],creationflags=DETACHED_PROCESS,close_fds=True) 
        return self.jsonify(code=0)

if __name__ == '__main__':
    pass
