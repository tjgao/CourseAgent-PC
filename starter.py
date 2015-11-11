import subprocess, sys
DETACHED_PROCESS = 8

print(sys.argv)

subprocess.Popen(["python","qrauth.py","abc"], creationflags=DETACHED_PROCESS, close_fds=True)

