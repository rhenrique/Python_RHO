import numpy as np
from PIL import ImageGrab
import cv2
import time

img1 = cv2.imread('Serpent_Spawn-0.png',0)
w, h = img1.shape[::-1]

def process_img(image):
    original_image = image
    # convert to gray
    processed_img = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # edge detection
    #processed_img =  cv2.Canny(processed_img, threshold1 = 200, threshold2=300)
    return processed_img

def main():
	last_time = time.time()
	while True:
		screen =  np.array(ImageGrab.grab(bbox=(0,40,800,640)))
		#print('Frame took {} seconds'.format(time.time()-last_time))
		last_time = time.time()
		new_screen = process_img(screen)
		res = cv2.matchTemplate(new_screen,img1,cv2.TM_CCOEFF_NORMED)
		threshold = 0.8
		loc = np.where( res >= threshold)
		for pt in zip(*loc[::-1]):
			cv2.rectangle(new_screen, pt, (pt[0] + w, pt[1] + h), (0,0,255), 2)
			print "reconhecido"
		cv2.imshow('window', new_screen)
        #cv2.imshow('window',cv2.cvtColor(screen, cv2.COLOR_BGR2RGB))
		if cv2.waitKey(25) & 0xFF == ord('q'):
			cv2.destroyAllWindows()
			break
			
main()
from IPython.display import Image
Image(filename='edge-detection.png') 