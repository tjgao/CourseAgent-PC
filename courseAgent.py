import qrauth, setting, singleinstance 
import cherrypy, multiprocessing, time, hashlib, sys
from uuid import getnode

class courseAgent:
    def __init__(self, st):
        self.users = []
        self.admin = None
        self.remote = st.getValue('server','remoteServer')
        try:
            self.port = int(st.getValue('global','server.socket_host'))
        except:
            self.port = 9503

    @cherrypy.expose
    def index(self):
        if self.admin is None:
            return "courseAgent is alive, but not connected."
        else:
            pass

    @cherrypy.expose
    def takeover(self, token):
        pass
        

def qrauth_proc(pipe, st, ev):
    m = hashlib.md5()
    m.update((str(getnode()) + str(time.time())).encode('utf-8'))
    code = m.hexdigest()
    q = qrauth.qrAuth(st)
    q.show(code)
    auth = q.getAuth()
    if auth is not None:
        ev.set()
        pipe.send(auth)
    else:
        pipe.close()


if __name__ == '__main__':
    single = singleinstance.singleInstance()
    if single.alreadyRunning():
        print('Hmm')
        sys.exit()
    s = setting.setting() 
    out_p, in_p= multiprocessing.Pipe()
    ev = multiprocessing.Event()

    authproc = multiprocessing.Process(target = qrauth_proc, 
            args = (in_p, s, ev,))
    authproc.start()
    authproc.join()

    res = None
    try:
        if ev.is_set():
            res = out_p.recv()
    except:
        pass
    out_p.close()
    in_p.close()
    if res is None:  sys.exit() 

    course = courseAgent(s)
    cherrypy.config.update({'server.socket_port':course.port})
    cherrypy.quickstart(course)

