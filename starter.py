import subprocess, sys
DETACHED_PROCESS = 8


subprocess.Popen(["python","main.py"], creationflags=DETACHED_PROCESS, close_fds=True)

