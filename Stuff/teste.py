import pyautogui
import time

time.sleep(3)
im = pyautogui.screenshot()
#screenWidth = pyautogui.locate('icon_tibia.png', im, grayscale=False)
pyautogui.moveTo(46,345)
pyautogui.click(clicks=2)
time.sleep(18)
pyautogui.press(['o','l','i','v','e','i','r','h'])
pyautogui.press('tab')
pyautogui.press(['1','2','1','3','2','9','3','1','2','4','1','4','r','a','@','#'])
pyautogui.press('enter')
pyautogui.hotkey('winleft','d')
pyautogui.moveTo(51,543)
pyautogui.click(clicks=2)
time.sleep(10)
authy = pyautogui.locateCenterOnScreen('authy.png')
pyautogui.click(authy)
time.sleep(1)
authy = pyautogui.locateCenterOnScreen('authy_button.png')
pyautogui.click(authy)
tibia = pyautogui.locateCenterOnScreen('tibia_mini.png')
pyautogui.click(tibia)
time.sleep(1)
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')





#print screenWidth



#while flecha != None:
#	pyautogui.press('subtract')
#	flecha = None
#	flecha = pyautogui.locateOnScreen('14.png')
	