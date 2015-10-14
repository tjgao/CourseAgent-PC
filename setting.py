import configparser, os

class setting:
    def __init__(self, target = None):
        if target is None: 
            target = os.path.dirname(os.path.realpath(__file__))
            target += (os.sep + 'settings.ini')
        self.config = configparser.RawConfigParser()
        self.config.read(target)

    def getValue(self, sec, key):
        if self.config is not None:
            return self.config.get(sec, key)

if __name__ == '__main__':
    s = setting()
    print( s.getValue('server','serverAddress'))
