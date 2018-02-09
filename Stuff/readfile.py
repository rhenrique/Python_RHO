import difflib
import sys

f = open('grupos.txt','r')

c = '034m'

b = []
num = 0
n = []
g = []

nome = raw_input("Digite o nome da pessoa:\n ")
n.append(nome)

def quest():
	grupo = raw_input("Digite o grupo da pessoa:\n ")
	g.append(grupo)

def ask():
    newline = raw_input("Adicionar mais usuarios?\n- [s] - para SIM;\n- [n] - para NAO.\n")
    return newline

quest()
yn = ask()
while (yn == 's'):
	quest()
	yn = ask()
	if (yn == 'n'):
		break
	else:
		True
		
login = str(n[num]) + '.txt'
w = open(login, 'w')

for line in g:
	grupos = g[num]
	w.write(grupos + '\n')
	num = num + 1

w.close()
small_file = open(login,'r')
long_file = open('grupos.txt','r')
output_file = open('output_file.txt','w')

try:
	small_lines = small_file.readlines()
	small_lines_cleaned = [line.rstrip().lower() for line in small_lines]
	long_file_lines = long_file.readlines()
	long_lines_cleaned = [line.rstrip().lower() for line in long_file_lines]
	
	for line in small_lines_cleaned:
		if line in long_lines_cleaned:
			output_file.writelines(line + '\n')
			
finally:
	small_file.close()
	long_file.close()
	output_file.close()
