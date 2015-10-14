import logging, os
from logging import handlers


class atLogger:
    def __init__(self):
        self.logger = logging.getLogger('Rotating Log')
        self.logger.setLevel(logging.INFO)
        self.fmt = logging.Formatter('%(asctime)s %(levelname)-7.7s %(message)s','%Y-%m-%d %H:%M:%S')
        self.setLogMode()

    def setDebug(self):
        self.logger.setLevel(logging.DEBUG)

    def addScreenMode(self):
        self.handler = logging.StreamHandler()
        self.handler.setFormatter(self.fmt)
        self.logger.addHandler(self.handler)

    def setLogMode(self):
        self.handler = handlers.RotatingFileHandler(
                os.path.dirname(os.path.realpath(__file__)) + os.sep + 'log' + os.sep + 'log.txt',
                maxBytes=(1<<20),
                backupCount=5)
        self.handler.setFormatter(self.fmt)
        self.logger.addHandler(self.handler)

    def info(self, msg):
        self.logger.info(msg)

    def debug(self, msg):
        self.logger.debug(msg)

    def warn(self, msg):
        self.logger.warn(msg)

    def error(self, msg):
        self.logger.error(msg)

    def critical(self, msg):
        self.logger.critical(msg)

    def exception(self, msg):
        self.logger.exception(msg)

logger = atLogger()
