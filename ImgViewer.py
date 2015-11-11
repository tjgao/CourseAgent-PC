from PIL import Image, ImageTk
import tkinter as tk
import ntpath, traceback, sys
import qrcode

class imgViewer:
    def __init__(self):
        self.viewer = None
        self.container = None
        self.running = True

    def openfile(self, target, title):
        im = Image.open(target)
        self.open(im, title)

    def getviewer(self):
        return self.viewer

    def exit(self):
        self.running = False

    def userhandler(self):
        pass

    def loop(self):
        if self.viewer is not None:
            try:
                while self.running:
                    self.userhandler()
                    self.viewer.update_idletasks()
                    self.viewer.update()
            except:
                pass

    def open(self, im, title):
        try:
            if self.viewer is None:
                self.viewer = tk.Tk()
            self.viewer.protocol('WM_CLOSE', self.close)
            self.viewer.attributes('-alpha', 0)
            self.viewer.lift()
            self.viewer.state('zoomed')
            self.viewer.update()
            win_width = self.viewer.winfo_screenwidth()
            win_height = self.viewer.winfo_screenheight()
            self.viewer.geometry(str(win_width) + 'x' + str(win_height))
            self.viewer.attributes('-alpha', 1)
            self.viewer.resizable(0,0)
            if self.container is None:
                self.container = tk.Canvas(self.viewer, bg = 'grey')
            self.container.pack(side='top', fill='both', expand=True)
            self.viewer.update()
            w, h = self.container.winfo_width(), self.container.winfo_height()
            iw, ih = im.size
            if iw > w or ih > h:
                if w*1.0/h > iw*1.0/ih:
                    ih = h
                    iw = h*iw*1.0/w
                else:
                    iw = w
                    ih = w*ih*1.0/h
                im.thumbnail((iw,ih),Image.ANTIALIAS)
            img = ImageTk.PhotoImage(im)
            self.container.image = img
            self.container.create_image(w/2, h/2, anchor='center', image = img, tags = 'bg_img')
            self.viewer.title(title)
            self.loop()
        except Exception as e: 
            print(e)
            pass

    def open2(self, code, title=''):
        qr = qrcode.QRCode(version=6)
        qr.add_data(code)
        qr.make()
        img = qr.make_image()
        self.open(img, title)

    def close(self):
        if self.container is not None:
            try:
                self.container.destroy()
            except:
                pass
        self.container = None
        if self.viewer is not None:
            try:
                self.viewer.destroy()
            except:
                pass
        self.viewer = None



if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2: 
        print('''
        Usage: 
            <program> -f filename title
            <program> -q string title
            ''')
        sys.exit()
    if sys.argv[1] in ['-f', '-q']:
        if sys.argv[1] == '-f':
            i = imgViewer()
            i.openfile(sys.argv[2],sys.argv[3])
        else:
            i = imgViewer()
            i.open2(sys.argv[2], sys.argv[3])

