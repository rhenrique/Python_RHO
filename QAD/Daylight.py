import pyautogui
import os
import time

time.sleep(5)
with open('daylight.txt') as f:
	dusers = f.readlines()
	for line in dusers:
		user = line.rstrip('\n')
		time.sleep(2)
		pyautogui.typewrite(user)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.press('enter')
		time.sleep(1)
		pyautogui.typewrite('GMT-3')
		time.sleep(1)
		pyautogui.press('f1')
		time.sleep(1)
		pyautogui.press('f1')
		time.sleep(1)
		pyautogui.press('f4')
		time.sleep(1)
		pyautogui.press('f4')
		time.sleep(1)
		pyautogui.press('f1')
		time.sleep(1)
	f.close()