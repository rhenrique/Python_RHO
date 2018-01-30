## COMPILAR: pyinstaller.exe --onefile Reg_Import.py

import subprocess
import glob
import psutil
import webbrowser
import time
import getpass
import os.path
import shutil
import sys
import re

i = 0
c_user = getpass.getuser()
java_path =  os.path.join('C:\\', 'Users',c_user, 'AppData', 'LocalLow', 'Sun', 'Java', 'Deployment', 'security')

#Verifica se ha Java 64-bit instalado
try:
	jre64 = os.listdir('C:\\Program Files\Java')
	for j_version in jre64:
		print j_version + "64-bit instalado. Por favor, remover antes de continuar."
	sys.exit()
except Exception:
	print "Java 64-bit nao instalado...."
	jre64 = 'aa'

try:
	j_version = os.listdir('C:\\Program Files (x86)\Java')
	for j_version in j_version:
		i = i + 1
		if i >= 2:
			print "Existem duas versoes do JAVA... remover uma das duas!"
			print ">>> " + j_version
			time.sleep(5)
			sys.exit()
except Exception:
	java = []
	java = glob.glob("jre*.exe")
	for file in java:
		print "Instalando Java >>> " + file
	process = subprocess.Popen([file + '/s', '/L', 'c:\pathsetup.log'],stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
	out, err = process.communicate()
	j_version = 'aa'
	time.sleep(5)
	
	
if j_version != 'jre6':
	try:
		subprocess.check_output(['javaws','-uninstall'])
	except Exception:
		print "Comando java nao esta sendo reconhecido pelo computador"
	os.system('cls') 
	print "ADICIONANDO SITES NA EXCECAO DO JAVA..."
	shutil.copy2('exception.sites', java_path)
	time.sleep(3)
else:
	print "JAVA 6 INSTALADO!!!."
	time.sleep(5)

PROCNAME = "iexplore.exe"

for proc in psutil.process_iter():
	# check whether the process name matches
	if proc.name() == PROCNAME:
		print u"Finalizando Internet Explorer"
		proc.kill()
		os.system('cls') 
	else:
		print u"Nao ha processo a ser finalizado"
		os.system('cls') 

arr = []
arr = glob.glob("*.reg")

for file in arr:
	print file
	subprocess.call(['reg','import',file])
	
ie = webbrowser.get(webbrowser.iexplore)
ie.open('http://imax.maxionwheels.com')
time.sleep(3)

