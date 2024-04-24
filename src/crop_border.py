#!/usb/bin/env python

import cv2
import numpy as np

def crop_border(image_path):

    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)


    # Apply adaptive thresholding to separate the paper from the background
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
    
    # Find contours of the paper
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Filter contours to find the one with the maximum area that satisfies size limits
    max_contour = None
    max_contour_area = 0

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        contour_area = cv2.contourArea(contour)
        if contour_area > max_contour_area:
            max_contour = contour
            max_contour_area = contour_area
    
    # Find the bounding rectangle of the paper contour
    x, y, w, h = cv2.boundingRect(max_contour)
    
    # Crop out the paper region including the dark border
    cropped_paper = img[y:y+h, x:x+w]

    # # AGAIN
    # gray = cv2.cvtColor(cropped_paper, cv2.COLOR_BGR2GRAY)


    # # Apply adaptive thresholding to separate the paper from the background
    # thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
    
    # # Find contours of the paper
    # contours, hierarchy = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # # Filter contours to find the one with the maximum area that satisfies size limits
    # max_contour = None
    # max_contour_area = 0
    # max_width = 3500
    # max_height = 3500
    # for contour in contours:
    #     x, y, w, h = cv2.boundingRect(contour)
    #     if w < max_width and h < max_height:
    #         contour_area = cv2.contourArea(contour)
    #         if contour_area > max_contour_area:
    #             max_contour = contour
    #             max_contour_area = contour_area
    
    # # Find the bounding rectangle of the paper contour
    # x, y, w, h = cv2.boundingRect(max_contour)
    
    # # Crop out the paper region including the dark border
    # cropped_paper = cropped_paper[y:y+h, x:x+w]
    # Image in grayscale
    cropped_paper_grayscale = cv2.cvtColor(cropped_paper, cv2.COLOR_BGR2GRAY)
    
    return cropped_paper_grayscale
