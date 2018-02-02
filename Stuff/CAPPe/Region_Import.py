## COMPILAR: pyinstaller.exe --onefile Region_Import.py
import subprocess
import glob
import re

arr = []
arr = glob.glob("C:\\Script_Cappe\*.reg")

# Importa todos os .reg na pasta RAIZ do Programa
for file in arr:
	print file
	subprocess.call(['reg','import',file])