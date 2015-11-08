import logging, os, json, atexit, shutil
import qrauth, singleinstance, courseAgent 
import cherrypy, multiprocessing, time, hashlib, sys, threading
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


def notifyExit(queue):
    queue.put_nowait(None)


if __name__ == '__main__':
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

    queue = multiprocessing.Queue(-1)
    atexit.register(notifyExit, queue)
    logging_listener = multiprocessing.Process(target = logging_proc, args=(queue,logging_configurer))
    logging_listener.daemon = True
    logging_listener.start()


    authproc = multiprocessing.Process(target = qrauth_proc, 
            args = (queue, worker_configurer, gconfig))
    authproc.start()
    authproc.join()

    
    logger = worker_configurer(queue)

    if gconfig.get('token') is None:
        logger.info('Window closed or auth failed, exit')
        sys.exit()

    course = courseAgent(gconfig)
    cherrypy.config.update({'server.socket_port':gconfig.get('port',9503)})
    cherrypy.quickstart(course)
