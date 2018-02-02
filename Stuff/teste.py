import pyautogui
import time
import datetime

now = datetime.datetime.now()
cdate = now.strftime("%Y-%m-%d %H-%M")


def principal():
	time.sleep(3)
	print ">>> Iniciando..."
	pyautogui.hotkey('winleft','d')
	pyautogui.moveTo(46,345)
	print ">>> Abrindo o jogo..."
	pyautogui.click(clicks=2)
	time.sleep(18)
	pyautogui.typewrite('LOGIN')
	pyautogui.press('tab')
	pyautogui.typewrite('SENHA AQUI')
	pyautogui.press('enter')
	time.sleep(5)
	pyautogui.press('h')
	pyautogui.press('enter')
	time.sleep(10)
	print ">>> Identificando se o jogo esta online..."
	on = pyautogui.locateCenterOnScreen('14.png')
	online(on)

def online(on):
	while on == None:
		time.sleep(10)
		print ">>> Nao esta online. Tentando identificar novamente..."
		on = locateCenterOnScreen('14.png')
	logout(on)
		
def logout(on):
	if on == None:
		online(on)
	else:
		print ">>> Se localizando dentro do jogo... primeira tentativa!"
		pyautogui.press('tab')
		pyautogui.moveTo(866,479)
		pyautogui.keyDown('shift'); 
		pyautogui.click(clicks=1)
		time.sleep(1)
		dist = pyautogui.locateCenterOnScreen('dist.png')
		if dist != None:
			print ">> IDENTIFICADO! - Efetuando o treinamento e logando..."
			pyautogui.keyUp('shift');
			pyautogui.click(button='right', clicks=1)
			time.sleep(2)
		else:
			print ">>> Se localizando dentro do jogo... segunda tentativa!"
			pyautogui.moveTo(798,464)
			pyautogui.keyDown('shift'); 
			pyautogui.click(clicks=1)
			time.sleep(1)
			dist = pyautogui.locateCenterOnScreen('dist.png')
		
		if dist != None:
			print ">> IDENTIFICADO! - Efetuando o treinamento e logando..."
			pyautogui.keyUp('shift');
			pyautogui.click(button='right', clicks=1)
			time.sleep(2)
		else:
			print ">>> Se localizando dentro do jogo... terceira e ultima tentativa!"
			pyautogui.moveTo(935,474)
			pyautogui.keyDown('shift'); 
			pyautogui.click(clicks=1)
			time.sleep(1)
			dist = pyautogui.locateCenterOnScreen('dist.png')
			
		if dist != None:
			pyautogui.keyUp('shift');
			pyautogui.click(button='right', clicks=1)
			time.sleep(2)
			print ">> IDENTIFICADO! - Efetuando o treinamento e logando..."
			print dist
			
		pyautogui.keyDown('shift');
		print ">>> Verificando se houve sucesso em treinar..."
		select = pyautogui.locateCenterOnScreen('tibia_select.png')
		if select != None:
			pyautogui.hotkey('alt','f4')
			print "Treinamento ocorreu sem problemas..."
			pyautogui.hotkey('alt','tab')
			time.sleep(2)
			pyautogui.screenshot(cdate + '.png')
		else:
			pyautogui.hotkey('ctrl','l')
			time.sleep(2)
			pyautogui.hotkey('alt','f4')
			print "Falha em fazer o treinamento... VERIFICAR"
			pyautogui.hotkey('alt','tab')
			time.sleep(2)
			pyautogui.screenshot(cdate + '.png')
			
			
principal()		
	 

	 




#print screenWidth



#while flecha != None:
#	pyautogui.press('subtract')
#	flecha = None
#	flecha = pyautogui.locateOnScreen('14.png')
	