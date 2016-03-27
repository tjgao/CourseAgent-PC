import logging, os, json, atexit, shutil, pythoncom
import qrauth, singleinstance, courseAgent 
import cherrypy, multiprocessing, time, hashlib, sys, threading, requests
from uuid import getnode
from QueueHandler import QueueHandler

logger = logging.getLogger(__name__)

def qrauth_proc(queue, configurer, gconfig):
    root = configurer(queue)
    m = hashlib.md5()
    m.update((str(getnode()) + str(time.time())).encode('utf-8'))
    code = m.hexdigest()
    q = qrauth.qrAuth(gconfig, root)
    q.show(code)



def logging_proc(queue, configurer):
    configurer() 
    while True:
        try:
            record = queue.get()
            if record is None: break
            logger = logging.getLogger(__name__)
            logger.handle(record)
        except(KeyboardInterrupt, SystemExit):
            raise

def logging_configurer():
    root = logging.getLogger(__name__)
    path =  os.path.dirname(os.path.realpath(__file__)) + os.sep + 'log'
    if not os.path.exists(path):
        os.mkdir(path)
    h = logging.handlers.RotatingFileHandler( path + os.sep + 'log.txt', 'a', 1024, 0)
    f = logging.Formatter('%(asctime)s %(levelname)-8s %(message)s','%Y-%m-%d %H:%M:%S')
    h.setFormatter(f)
    root.addHandler(h)
    return root


def worker_configurer(queue):
    h = QueueHandler(queue)
    root = logging.getLogger(__name__)
    root.addHandler(h)
    root.setLevel(logging.DEBUG)
    return root


def getserver():
    url = 'https://raw.githubusercontent.com/tjgao/SpringBoard/master/server.json'
    r = requests.get(url)
    try:
        j = r.json()
        return j.get('server')
    except:
        print('should not happen');
        pass

def notifyExit(queue):
    queue.put_nowait(None)


def onThreadStart(threadIndex):
  pythoncom.CoInitialize()


def cleardir(directory):
    fl = os.listdir(directory)
    for f in fl:
        try:
            os.remove(directory + os.sep + f)
        except:
            shutil.rmtree(directory + os.sep + f)

if __name__ == '__main__':
    #DEBUG = True 
    DEBUG = False 
    single = singleinstance.singleInstance()
    if single.alreadyRunning():
        sys.exit()

    mgr = multiprocessing.Manager()
    gconfig = mgr.dict()
     
    cfgjson = 'config.json'
    if os.path.exists(cfgjson):
        with open(cfgjson) as f:
            dd = json.loads(f.read())
            gconfig.update(dd)

    tmp = gconfig.get('server')
    if tmp is None or len(tmp) == 0:
        gconfig['server'] = getserver()

    queue = multiprocessing.Queue(-1)

    atexit.register(notifyExit, queue)

    logging_listener = multiprocessing.Process(target = logging_proc, args=(queue,logging_configurer))
    logging_listener.daemon = True
    logging_listener.start()

    course = courseAgent.courseAgent(gconfig, queue, worker_configurer)
    filedir = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__))),'files')
    if not os.path.exists(filedir):
        os.mkdir(filedir)
    else:
        cleardir(filedir)

    if not DEBUG:
        authproc = multiprocessing.Process(target = qrauth_proc, 
            args = (queue, worker_configurer, gconfig))
        authproc.start()
        authproc.join()
    else:
        gconfig['token'] = 'abcdef'
        gconfig['token2'] = '123456'

    
    logger = worker_configurer(queue)

    if gconfig.get('token') is None:
        logger.info('Window closed or auth failed, exit')
        sys.exit()


    cherrypy.engine.subscribe('start_thread', onThreadStart)

    cherrypy.config.update({
        'server.thread_pool': 40,
        'server.thread_pool_max': -1,
        'server.socket_host': '0.0.0.0',
        'server.socket_port':gconfig.get('port',9503)
        #        'tools.sessions.on':True,
        #        'tools.sessions.locking':'explicit',
    })
    conf = {
        '/files':{
            'tools.staticdir.on':True,
            'tools.staticdir.dir':filedir
        }
    }
    cherrypy.quickstart(course, config=conf)

