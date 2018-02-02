import pyautogui
import time
import sys

n = []
l = []
a = []

num = 0

def quest():
	nome = raw_input("Digite o nome da pessoa:\n ")
	n.append(nome)
	login = raw_input("Digite o login da pessoa:\n ")
	l.append(login)
	amb = raw_input("Escolha um ambiente:\n - [128] - ACO;\n - [130] - Aluminio.\n")
	a.append(amb)

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
	



#note = pyautogui.locateCenterOnScreen('imagens/notepad.png')
#print note
#pyautogui.click(note)
#note = pyautogui.locateCenterOnScreen('imagens/notepad_1.png')
#print note

#Primeira linha
#pyautogui.moveTo(note)
#pyautogui.moveRel(-42, 40)
#pyautogui.click()
#pyautogui.dragRel(50, 0, 2, button='left')
#pyautogui.hotkey('ctrl', 'c')
pyautogui.hotkey('winleft','d')

#Abrindo o QAD
pyautogui.press(['Q','A','D','-','P','R','O','D'])
pyautogui.press('enter')

#Logando no QAD
time.sleep(4)
pyautogui.typewrite('oliveirh')
pyautogui.press('enter')
time.sleep(1)
pyautogui.typewrite('312414ra@#')  
pyautogui.press('enter')
time.sleep(3)
pyautogui.press('1')
pyautogui.press('enter')

for item in n:
	time.sleep(2)
	if a[num] == '128':
		pyautogui.press('1')
	else:
		pyautogui.press('2')
	pyautogui.press('enter')
	time.sleep(3)
	pyautogui.press('enter')
	time.sleep(3)
	pyautogui.press('enter')
	time.sleep(3)
	pyautogui.typewrite('36.3.1')
	pyautogui.press('enter')
	time.sleep(3)
	pyautogui.typewrite(l[num])
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.typewrite(n[num])
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.typewrite('po')
	pyautogui.press(['enter','enter'])
	time.sleep(1)
	pyautogui.typewrite('brs')
	pyautogui.press(['enter','enter','enter'])
	time.sleep(1)
	pyautogui.typewrite('PRIMARY')
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.typewrite('GMT-2')
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.typewrite('mailx')
	pyautogui.press(['enter','enter','enter'])
	time.sleep(1)
	pyautogui.typewrite('s')
	pyautogui.press(['enter','enter','enter'])
	time.sleep(1)
	pyautogui.typewrite('QAD_DEF')
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.press(['enter','enter','enter'])
	time.sleep(2)
	pyautogui.typewrite(a[num])
	pyautogui.press('enter')
	pyautogui.press(['enter','enter'])
	pyautogui.typewrite('s')
	pyautogui.press('enter')
	pyautogui.press('f4')
	pyautogui.press('f4')
	pyautogui.typewrite('s')
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.press('f4')
	pyautogui.press('enter')
	time.sleep(1)
	pyautogui.hotkey('ctrl', 'c')
	time.sleep(1)
	print num
	num = num + 1
	
	







	
	
	
	
	




#pyautogui.press('u')
#pyautogui.press('enter')
#time.sleep(1)
#pyautogui.typewrite('sudo passwd oliveirh')
#pyautogui.press('enter')
#time.sleep(1)
#pyautogui.typewrite('312414ra@#') 
#pyautogui.press('enter')
#time.sleep(1)
#pyautogui.typewrite('312414ra@#') 
#pyautogui.press('enter')
#time.sleep(1)
#pyautogui.typewrite('312414ra@#')
#pyautogui.press('enter')
#time.sleep(1)
#pyautogui.typewrite('exit')
#pyautogui.press('enter') 





#Segunda linha
#pyautogui.moveTo(data)
#pyautogui.moveRel(-42, 60)
#pyautogui.dragRel(250, 0, 2, button='left')