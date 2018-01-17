import subprocess
import glob
import psutil
import webbrowser
import time
import getpass
import os.path
import shutil

c_user = getpass.getuser()
java_path =  os.path.join('C:\\', 'Users',c_user, 'AppData', 'LocalLow', 'Sun', 'Java', 'Deployment', 'security')
try:
	subprocess.check_output(['javaws','-uninstall'])
except Exception:
	print "Possivelmente nao ha JAVA instalado, favor verificar"

shutil.copy2('exception.sites', java_path)

PROCNAME = "iexplore.exe"

for proc in psutil.process_iter():
	# check whether the process name matches
	if proc.name() == PROCNAME:
		print u"Finalizando Internet Explorer"
		proc.kill()
	else:
		print u"Nao ha processo a ser finalizado"

arr = []
arr = glob.glob("*.reg")

for file in arr:
	print file
	subprocess.call(['reg','import',file])
	
ie = webbrowser.get(webbrowser.iexplore)
ie.open('http://imax.maxionwheels.com')
time.sleep(3)

