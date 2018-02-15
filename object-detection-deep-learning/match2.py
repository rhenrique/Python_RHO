import numpy as np
from PIL import ImageGrab
import cv2
import time
import pyautogui


img1 = cv2.imread('images/foto.jpg',0)
img2 = cv2.imread('images/foto2.jpg',0)
img3 = cv2.imread('images/store3.jpg',0)
w, h = img1.shape[::-1]

def process_img(image):
    original_image = image
    # convert to gray
    processed_img = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # edge detection
    #processed_img =  cv2.Canny(processed_img, threshold1 = 200, threshold2=300)
    return processed_img

def check(pt):
	
	print "reconhecido 2"
	print pt
	while True:
		print "reconhecido 2"
		screen =  np.array(ImageGrab.grab(bbox=(0,40,800,640)))
		#print('Frame took {} seconds'.format(time.time()-last_time))
		last_time = time.time()
		new_screen = process_img(screen)
		res = cv2.matchTemplate(new_screen,img2,cv2.TM_CCOEFF_NORMED)
		min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
		threshold = 0.8
		loc = np.where( res >= threshold)
		print loc
		top_left = max_loc
		bottom_right = (top_left[0] + w, top_left[1] + h)
		#for pt in zip(*loc[::-1]):
		#cv2.rectangle(new_screen, pt, (pt[0] + w, pt[1] + h), (0,0,255), 2)
		pt = []
		for pt in zip(*loc[::-1]):
			cv2.rectangle(new_screen, top_left, bottom_right, (0,0,255), 2)
			if pt != None:
				print pt
				print 'entrou aqui?'
			else:
				return pt
			#lala = check(pt)
			#utogui.moveTo(pt)
			
		cv2.imshow('window2', new_screen)
        #cv2.imshow('window',cv2.cvtColor(screen, cv2.COLOR_BGR2RGB))
		if cv2.waitKey(25) & 0xFF == ord('q'):
			cv2.destroyAllWindows()
			break
		return pt
		
def walk():
	if 	

def main():
	last_time = time.time()
	while True:
		screen =  np.array(ImageGrab.grab(bbox=(0,40,800,640)))
		#print('Frame took {} seconds'.format(time.time()-last_time))
		last_time = time.time()
		new_screen = process_img(screen)
		res = cv2.matchTemplate(new_screen,img1,cv2.TM_CCOEFF_NORMED)
		min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
		threshold = 0.8
		loc = np.where( res >= threshold)
		print loc
		top_left = max_loc
		bottom_right = (top_left[0] + w, top_left[1] + h)
		#for pt in zip(*loc[::-1]):
		#cv2.rectangle(new_screen, pt, (pt[0] + w, pt[1] + h), (0,0,255), 2)
		for pt in zip(*loc[::-1]):
			cv2.rectangle(new_screen, top_left, bottom_right, (0,0,255), 2)
			print pt
			lala = check(pt)
			#utogui.moveTo(pt)
			
		cv2.imshow('window', new_screen)
        #cv2.imshow('window',cv2.cvtColor(screen, cv2.COLOR_BGR2RGB))
		if cv2.waitKey(25) & 0xFF == ord('q'):
			cv2.destroyAllWindows()
			break
		
			
main()