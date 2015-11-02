import poplib
import win32api
import win32process
import time
email = '996825244@qq.com'
host = 'pop.qq.com'
port = '995'
try:
    f = open('config.ini','r')
except IOError:
    print "open config failed"
password = f.read().split()[1]
f.close()

server = poplib.POP3_SSL(host)
server.set_debuglevel(0)
print server.welcome
server.user(email)
server.pass_(password)

print('Messages: %s. Size: %s' % server.stat())
resp, mails, octets = server.list()
print(mails)
index = len(mails)
resp, lines, octets = server.retr(index)
for mycontent in lines:
    if 'Subject'.lower() in mycontent.lower():
        break
# print type(msg)

print mycontent
content = mycontent.split(':')[1]
content1 = content.strip()

if content1 == 'test':
    hand = win32process.CreateProcess('C:\Program Files\Netease\CloudMusic\cloudmusic.exe',
                                      '',None,None,0,win32process.CREATE_NO_WINDOW,None,None,
                                      win32process.STARTUPINFO())
    time.sleep(3)
    hand1 = win32api.ShellExecute(0,'open','shinian.mp3','','',1)
wait1 = raw_input("wait............................:")
if wait1 == 'close' or wait1 == 'c':
    win32process.TerminateProcess(hand[0],0)
    # win32api.CloseHandle(hand)
server.quit()
