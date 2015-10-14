import win32serviceutil, win32service, win32event, servicemanager, time, logging
import cherrypy
import courseAgent

logging.basicConfig(
    filename = 'service.log',
    level = logging.DEBUG,
    format = '[courseAgent-service] %(levelname) - 7.7s %(message)s'
)


class courseAgentSvc(win32serviceutil.ServiceFramework):
    _svc_name = 'courseAgent-Service'
    _svc_display_name_ = 'courseAgent Service'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)

    def SvcDoRun(self):
        cherrypy.tree.mount(courseAgent(), '/')
        chrrypy.config.update({
            'global':{
                'engine.autoreload.on': False,
                'log.screen': False,
                'engine.SIGHUP': None,
                'engine.SIGTERM': None
            }
        })
        cherrypy.server.quickstart()
        cherrypy.engine.start(blocking=False)
        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        cherry.server.stop()
        win32event.SetEvent(self.stop_event)
