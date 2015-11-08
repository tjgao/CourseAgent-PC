import subprocess
DETACHED_PROCESS = 8

subprocess.Popen(["python","qrauth.py","abc"], creationflags=DETACHED_PROCESS, close_fds=True)

