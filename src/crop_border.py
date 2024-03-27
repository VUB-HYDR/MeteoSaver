#!/usb/bin/env python

import cv2
import numpy as np

img = cv2.imread('sofwin.png')
gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
_,thresh = cv2.threshold(gray,1,255,cv2.THRESH_BINARY)

contours,hierarchy = cv2.findContours(thresh,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
cnt = contours[0]
x,y,w,h = cv2.boundingRect(cnt)

crop = img[y:y+h,x:x+w]
cv2.imwrite('sofwinres.png',crop)

