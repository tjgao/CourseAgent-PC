import win32com.client

sh = win32com.client.Dispatch('Shell.Application')
wnds = sh.Windows()
try:
    for i in range(wnds.Count):
        print(i)
        print(sh[i].LocationName)
        print(sh[i].LocationURL)
except:
    pass
