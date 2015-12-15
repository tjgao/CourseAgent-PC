import logging, os, json, atexit, shutil
import cherrypy, multiprocessing, time, hashlib, sys, enum, requests
import ImgViewer, win32gui, win32con, logging, subprocess, time, datetime
import threading, queue
from uuid import getnode
from PIL import Image
from queue import Queue
from MSPPT import MSPPT
from win import win

DETACHED_PROCESS = 8

logger = logging.getLogger(__name__)


viewer_title = '__IMAGE_VIEWER__'
#viewer_title = '_._ 图片查看 _._'
attend_title = '_._ 课堂签到 _._'

#class QTYPE(enum.Enum):
#    TIMEOUT = 0
#    CHATMSG = 1
#    SHOWUP = 2
#    SHOWOFF = 3
#    RAISEHAND = 4
#    GAMEOVER = 1000
#
#class qulet:
#    def __init__(self, qtype, data = None):
#        self.qtype = qtype
#        self.data = data
#        self.code = 0
#    def __str__(self):
#        return json.dumps(self.__dict__)


class courseAgent:
    def __init__(self, gconfig, queue, configurer):
        self._curppt = None
        self._curpic = None
        self.ppt = MSPPT()
        self._user = {}
        self.gconfig = gconfig
        self.longpoll = self.gconfig.get('longpoll',10)
        self.queue = queue
        self.configurer = configurer
        self.logger = logging.getLogger(__name__)
        #self.messages = []
        self.lock = threading.Lock()
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        if not os.path.exists(path):
            os.mkdir(path)

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
        img = Image.open(src)
        img.thumbnail((width,height))
        img.save(dst)        
        return dst

    def reduceSize(self, src):
        width, height = 1280, 1024
        img = Image.open(src)
        img.thumbnail((width, height))
        img.save(src)

    def jsonify(self, **args):
        return json.dumps(args)


    def _closewindow(self, title):
        found = False
        try:
            w = win32gui.FindWindow(None, title)
            while w:
                found = True
                win32gui.PostMessage(w, win32con.WM_CLOSE, 0, 0)
                w = win32gui.FindWindow(None, title)
        except:
            pass
        finally:
            return found

#    def notifyChange(self, fn):
#        d = dict()
#        d['filename'] = fn
#        item = qulet(QTYPE.SHOWUP, d)
#        for i in self._user:
#            self._user[i].get('queue').put_nowait(item)

    def checktoken(self):
        #return 1,2
        uid = cherrypy.request.headers.get('id')
        tok = cherrypy.request.headers.get('token')
        #print(uid,tok)
        t1, t2 = self.gconfig.get('token'), self.gconfig.get('token2')
        #print(t1,t2)
        if uid is None: return None
        if tok not in (t1, t2): return None
        if self._user.get(uid) is None: 
            self._user[uid] = {} 
            self._user[uid]['uid'] = uid
            self._user[uid]['queue'] = Queue()
            self._user[uid]['rank'] = (tok == t1) and 2 or 1
        #print( uid, self._user[uid]['rank'])
        return uid, self._user[uid]['rank']

    def _openpic(self, filename=None):
        self._closewindow(viewer_title) 
        self._closewindow(attend_title)
        if filename:
            self._curpic = filename
        if self._curpic is None: return
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        subprocess.Popen(['python','ImgViewer.py','-f',path + os.sep + self._curpic,viewer_title],creationflags=DETACHED_PROCESS,close_fds=True) 
        time.sleep(1)
        self._broadcast()

    def _openppt(self, filename=None):
        if filename :
            self._curppt = filename
        if self._curppt is None: return
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        self.ppt.open(path + os.sep + self._curppt)
        # slide's name is 1.png,2.png,....
        self.ppt.makepics( path ) 
        # manipulate ppt window
        pptwin = win()
        pptwin.wildfind(self._curppt)
        n = 3
        while not pptwin.valid() and n:
            time.sleep(0.3)
            pptwin.wildfind(self._curppt)
            n = n - 1
        if pptwin.valid():
            pptwin.foreground()
            pptwin.activate()
        self._broadcast()



    def _whattoshow(self):
        if self._curpic:
            picwin = win()
            picwin.wildfind(viewer_title)
            #picwin.wildfind('Notepad')
            if picwin.valid():
                print('I think we can find the window')
                return self._curpic
        if self._curppt:
            if self.ppt and self.ppt.opened():
                return str(self.ppt.curslide()) + '.png'
        return ''

    def _broadcast(self):
        print('before that, we have ' + str(self._curpic) + ' : ' + str(self._curppt))
        show = self._whattoshow()
        print('Be about to broadcast: ' + show)
        # make sure everybody sees it
        for i in self._user:
            q = self._user[i].get('blackboard')
            if q: q.put_nowait(show)

    @cherrypy.expose
    def toggleppt(self):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        try:
            if self.ppt and self.ppt.opened():
                self.ppt.close()
            else:
                self._openppt()
        except Exception as e:
            print(e)
        return self.jsonify(code=0)

    @cherrypy.expose
    def togglepic(self):
        a = self._closewindow(viewer_title) 
        b = self._closewindow(attend_title)
        if a or b:
            return self.jsonify(code=0)
        self._openpic()
        return self.jsonify(code=0)


    @cherrypy.expose
    def index(self):
        return 'I am alive, ready to go.'

    @cherrypy.expose
    def alive(self):
        return self.jsonify(code=0)


    @cherrypy.expose
    def gameover(self):
        #        item = qulet(QTYPE.GAMEOVER)
        #        for i in self._user:
        #            self._user[i].get('queue').put_nowait(item)
        #        time.sleep(1)
        print('For test, I am not really shutting down')
        return self.jsonify(code=0)
        cherrypy.engine.exit()
        os.system('shutdown /s /t 0')
        return self.jsonify(code=0)

    @cherrypy.expose
    def register(self, qtype, uname, nickname, headimg = None):
        uid = cherrypy.request.headers.get('id')
        tok = cherrypy.request.headers.get('token')
        rank = 0
        if tok == self.gconfig.get('token'): rank = 2
        elif tok == self.gconfig.get('token2'): rank = 1
        else : return self.jsonify(code=1, msg='认证失败！')
        self._user[uid] = {} 
        self._user[uid]['uid'] = uid
        self._user[uid]['uname'] = uname
        self._user[uid]['nickname'] = nickname
        self._user[uid]['headimg'] = headimg
        if qtype == 'blackboard':
            self._user[uid]['blackboard'] = Queue(0)
        elif qtype == 'chatroom':
            self._user[uid]['chatroom'] = Queue(0)
        else:
            return self.jsonify(code=1,msg='参数错误！')
        self._user[uid]['rank'] = rank
        return self.jsonify(code=0)

    
    @cherrypy.expose
    def unregister(self):
        uid = cherrypy.request.headers.get('id')
        tok = cherrypy.request.headers.get('token')
        try:
            del(self._user[uid])
        except:
            pass
        return self.jsonify(code=0)


    @cherrypy.expose
    def listen(self):
        ret = self.checktoken() 
        if ret is None :
            return self.jsonify(code=1, msg='No permission')
        uid = ret[0]
        q = self._user[uid].get('chatroom')
        try:
            m = q.get(True, 30)
            return self.jsonify(code=0, data=m)
        except queue.Empty:
            return self.jsonify(code=0, data=None)

#        no = self._user[uid].get('msgNo')
#        if no is None: no = 0
#        l = len(self.messages)
#        if no == l: return self.jsonify(code=0,data=[])
#        self._user[uid]['msgNo'] = l
#        if l - no > 100:
#            return self.jsonify(code=0, data=self.messages[l-100:])
#        else:
#            return self.jsonify(code=0, data=self.messages[l-no:]);



    @cherrypy.expose
    def chat(self, msg):
        ret = self.checktoken() 
        if ret is None :
            return self.jsonify(code=1, msg='No permission')
        uid = ret[0]
        if self._user[uid].get('chatroom') is None:
            return self.jsonify(code=1, msg="Not registered")
        try:
            m = { 'text':msg, 'headimg':self._user[uid].get('headimg'),
                    'realname':self._user[uid].get('uname'),
                    'time':str(datetime.datetime.now()), 'uid':uid}
            for x in self._user:
                q = self._user[x].get('chatroom')
                if q:
                    q.put_nowait(m)
            return self.jsonify(code=0)
        except Exception as e:
            print(e)
            return self.jsonify(code=1, msg="Exception happens when inserting")

    @cherrypy.expose
    def fastwatch(self):
        ret = self.checktoken() 
        if ret is None :
            return self.jsonify(code=1, msg='No permission')
        if self._curpic:
            picwin = win()
            picwin.wildfind(viewer_title)
            if picwin.valid():
                return self.jsonify(code=0, data=self._curpic)
        if self.ppt and self.ppt.opened():
            return self.jsonify(code=0, data=str(self.ppt.curslide())+'.png')   
        return self.jsonify(code=0, data='')


    @cherrypy.expose
    def watch(self):
        ret = self.checktoken() 
        if ret is None :
            return self.jsonify(code=1, msg='No permission')
        uid = ret[0]
        q = self._user[uid].get('blackboard')
        if not q: return self.jsonify(code=1, msg='Not registered')
        #if there is a pic there, choose to show it first
        try:
            m = q.get(True, 30)
            return self.jsonify(code=0, data=m)
        except queue.Empty:
            return self.jsonify(code=0, data=None)


    @cherrypy.expose
    def currentslide(self):
        ret = self.checktoken()
        if ret is None: return self.jsonify(code=1)
        if self.ppt and self.ppt.opened():
            return self.jsonify(idx=self.ppt.curslide(), code=0)
        return self.jsonify(code=1, msg='Powerpoint is not running')

    @cherrypy.expose
    def raisehand(self, filename):
        pass

#    @cherrypy.expose
#    def closeviewer(self):
#        ret = self.checktoken() 
#        if ret is None or ret[1] < 2:
#            return self.jsonify(code=1, msg='No permission')
#        self._closewindow(viewer_title)
#        self._closewindow(attend_title)
#        return self.jsonify(code=0)
        
    @cherrypy.expose
    def toggleattend(self, attendstr):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        self._closewindow(viewer_title)
        if self._closewindow(attend_title):
            return self.jsonify(code=0)
        subprocess.Popen(['python','ImgViewer.py','-q',attendstr,attend_title],creationflags=DETACHED_PROCESS,close_fds=True) 
        return self.jsonify(code=0)

    @cherrypy.expose
    def permitshow(self, filename):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        if filename.endswith('.pptx') or filename.endswith('.ppt'):
            self._curppt = path + os.sep + filename
        else:
            self._curpic = path + os.sep + filename
        return self.jsonify(code=0)

    @cherrypy.expose
    def upload(self, file, name, ext):
        ret = self.checktoken() 
        if ret is None:
            return self.jsonify(code=1, msg='No permission')

        path = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        m = hashlib.md5()
        m.update((name + str(time.time())).encode('utf-8'))
        dest = path + os.sep + m.hexdigest() + ext.lower()
        with open(dest, 'wb') as f:
            while True:
                data = file.file.read(8192)
                if not data: break
                f.write(data)
        if ext.lower() in ['.jpg','.png','.gif','.bmp','.jpeg']:
            thumb = self.thumb(dest)
            self.reduceSize(dest)
            if ret[1] >= 2:
                self._curpic = os.path.basename(dest)
                self._openpic()
            return self.jsonify(thumb=os.path.basename(thumb), filename=os.path.basename(dest), code=0)
        elif ext.lower() in ['.ppt','.pptx']:
            if ret[1] < 2: return self.jsonify(code=1)
            self._curppt = os.path.basename(dest)
            self._openppt()
            if not self.ppt or not self.ppt.opened():
                return self.jsonify(code=1, msg='打开PPT失败！')
            return self.jsonify(filename = os.path.basename(dest), code=0)
        else:
            return self.jsonify(code=1, msg='只能上传图片或者Powerpoint文件！')

    @cherrypy.expose
    def downloadppt(self, url):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        fn = url.split('/')[-1]
        path =  os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
        fpath = path + os.sep + fn 
        url = self.gconfig.get('server') + '/' + url 
        r = requests.get(url, stream=True)
        if r.status_code != 200:
            return self.jsonify(code=1, msg="无法从指定地址下载！")
        with open(fpath, 'wb') as f:
            r.raw.decode_content = True
            shutil.copyfileobj(r.raw,f)
            #for chunk in r.iter_content(chunk_size=1024):
            #    if chunk: f.write(chunk)
        self._openppt(fn)
        if not self.ppt or not self.ppt.opened():
            return self.jsonify(code=1, msg='打开PPT失败！')
        return self.jsonify(code=0)

    @cherrypy.expose
    def openppt(self, filename = None):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt:
            if self.ppt.opened():
                self.ppt.close()
        else:
            self.ppt = MSPPT()
        if filename is None and self._curppt is None:
            return self.jsonify(code=1, msg='尚未指定Powerpoint文件！')
        self._openppt(filename)
        if not self.ppt or not self.ppt.opened():
            return self.jsonify(code=1, msg='打开PPT失败！')
        return self.jsonify(code=0) 

    @cherrypy.expose
    def ppt_next(self):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')

        if self.ppt and self.ppt.opened():
            self.ppt.next()
            self._broadcast()
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_prev(self):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')

        if self.ppt and self.ppt.opened():
            self.ppt.prev()
            self._broadcast()
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_pages(self):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            return self.jsonify(code=0, data=[self.ppt.getpages(),self.ppt.curslide()])
        return self.jsonify(code=1, msg='PPT is not running')

    @cherrypy.expose
    def ppt_goto(self, page):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.goto(page)
            self._broadcast()
        return self.jsonify(code=0)


    @cherrypy.expose
    def ppt_first(self):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.first()
            self._broadcast()
        return self.jsonify(code=0)

    @cherrypy.expose
    def ppt_last(self):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')
        if self.ppt and self.ppt.opened():
            self.ppt.last()
            self._broadcast()
        return self.jsonify(code=0)
        
    @cherrypy.expose
    def show(self, filename):
        ret = self.checktoken() 
        if ret is None or ret[1] < 2:
            return self.jsonify(code=1, msg='No permission')

        # start a image viewer to open the picture 
        self._openpic(filename)
        return self.jsonify(code=0)

if __name__ == '__main__':
    pass
