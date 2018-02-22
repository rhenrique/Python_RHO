import numpy as np
import cv2

watch_cascade = cv2.CascadeClassifier('cascade.xml')


cap = cv2.VideoCapture(0)

while 1:
	ret, img = cap.read()
	gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    
    # add this
    # image, reject levels level weights.
	watches = watch_cascade.detectMultiScale(gray, 50, 50)
    
    # add this
	print('entrou aqui')
	for (x,y,w,h) in watches:
		cv2.rectangle(img,(x,y),(x+w,y+h),(255,255,0),2)
		print('entrou aqui 2')



	cv2.imshow('img',img)
	k = cv2.waitKey(30) & 0xff
	if k == 27:
		break

cap.release()
cv2.destroyAllWindows()