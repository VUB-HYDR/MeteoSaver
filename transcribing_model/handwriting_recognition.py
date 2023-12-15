#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import numpy as np
## import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from paddleocr import PaddleOCR,draw_ocr
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment


## Function for preprocessing image from original color to black&white image using binarization
def preprocessing(image_path):

    ## Convert the original image to grayscale and then carry out binarization
    image_grayscale = cv2.imread(image_path, cv2.COLOR_BGR2GRAY)
    image_binarized = cv2.adaptiveThreshold(test_image_before_preprocessing_grayscale, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 91,6) 

    ## Here, the value 91 in the cv.adaptiveThreshold function call represents the size of the neighborhood area used for adaptive thresholding. It defines the size of the pixel neighborhood around each pixel used to calculate the threshold value for that pixel.
    ## Note that this value can be adjusted depending on your preference
    ## Here's what it means:
    ## In adaptive thresholding, the threshold value is calculated for each pixel based on the intensity of its surrounding neighborhood. The 91 specifies the size of the neighborhood window.
    ## Specifically, for each pixel in the image, a local threshold is calculated by taking the mean (or weighted mean) of the pixel values in a neighborhood window of size 91x91 (i.e., 91 pixels in width and 91 pixels in height) centered on that pixel.
    ## The calculated local threshold is then used to classify the pixel as either foreground (white) or background (black) based on whether the pixel's intensity is greater or less than the local threshold.
    ## You can adjust the size of the neighborhood window by changing the value to a different odd number. A larger neighborhood size can make the thresholding process more adaptive to larger variations in pixel intensity but may also blur small details.

    ## In the context of the OpenCV function cv.adaptiveThreshold, the value 255 represents the maximum pixel value for the output image, which is used to represent white in a binary image.
    ## Here's what it means in more detail:
    ## In binary images (typically used for thresholding), you have two values: 0 and a maximum value (usually 255), which correspond to black and white, respectively.
    ## When you set the thresholding method to cv.ADAPTIVE_THRESH_GAUSSIAN_C and specify cv.THRESH_BINARY, you are creating a binary image where pixels in the source image that are above the calculated threshold become white (255), and pixels below the threshold become black (0).

    ## To make the blacks darker in your processed image, you can adjust the thresholding parameters of the cv.adaptiveThreshold function. Specifically, you can try increasing the C parameter (here the C = 6), which controls the constant subtracted from the mean or weighted mean. Increasing C will make the thresholding result darker because it lowers the threshold level.

    return image_binarized

def table_detection_model(image_binarized):

    ## Here, we import the binarized image on which the table detection will be undertaken
    image = cv2.imread(image_binarized, cv2.IMREAD_GRAYSCALE)
    ## We, then produce an inverted image with global thresholding.
    ## Here, the threshold value for pixel intensities = 128, and the value 255 is assigned if the pixel value is above the threshold
    thresh,img_bin = cv2.threshold(image,128,255,cv2.THRESH_BINARY)  ## The binary thesholding method is used here
    img_bin = 255-img_bin
    
    ## Detect the vertical lines in the image
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, np.array(image).shape[1]//150))
    eroded_image = cv2.erode(img_bin, vertical_kernel, iterations=1)
    vertical_lines = cv2.dilate(eroded_image, vertical_kernel, iterations=1)

    ## Detect the horizontal lines in the image
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(image).shape[1]//150, 1))
    eroded_image= cv2.erode(img_bin, hor_kernel, iterations=1)
    horizontal_lines = cv2.dilate(eroded_image, hor_kernel, iterations=1)

    ## Blending the imaegs with the vertical lines and the horizontal lines 
    vertical_horizontal_lines = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
    vertical_horizontal_lines = cv2.erode(~vertical_horizontal_lines, hor_kernel, iterations=2)
    
    ## Identifing the cells
    gray_image = cv2.cvtColor(vertical_horizontal_lines, cv2.COLOR_BGR2GRAY)
    contours, hierarchy = cv2.findContours(gray_image, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    ## Create an Excel workbook and a worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = 'OCR_Results'
    ## Folder to save excel file outputs
    save_folder = './output'

    boxes = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        
        ## Calculate a factor to increase the bounding box area (e.g., 80% larger)
        increase_factor = 0.8  # Modify this factor as needed
        
        x -= int(w * increase_factor)  # Increase width of bounding box
        ##y -= int(h * increase_factor)  # Increase height of bounding box, IF NECESSARY. Here, I did not do so
        w += int(2 * w * increase_factor)  # Increase width of bounding box
        ##h += int(2 * h * increase_factor)  # Increase height of bounding box, IF NECESSARY. Here, I did not do so
        
        if (w<1000 and h<500):
            
            image_with_contours = cv2.rectangle(image,(x,y),(x+w,y+h),(0,255,0),2)
            boxes.append([x,y,w,h])

            ## Crop each cell using the bounding rectangle coordinates
            cropped_image = image[y:y+h, x:x+w]

            ## Perform Optical Character recognition on the cropped cells
            ocr_result = ocr.ocr(cropped_image, cls = True)

            ## Transfer reconized characters to Excel file with co-ordinates of the cropped image from the original image
            if ocr_result is not None and ocr_result[0] is not None:
            cell_values = ocr_result[0] ## The index [0] is for ocr results and the index [1] is for the confidence ratio
            text_in_cells = [line[1][0] for line in cell_values]

            for text in range(len(text_in_cells)):
                value =  text_in_cells[text]
                cell_column = openpyxl.utils.get_column_letter(x//100)  # Assuming x-coordinate translates to columns
                cell_row = y//100  # Assuming y-coordinate translates to rows
                
                # Convert the column letter to its corresponding index
                col_index = openpyxl.utils.column_index_from_string(cell_column)
                
                # Write the OCR value to the cell in the Excel sheet
                cell = ws.cell(row=int(cell_row) + text, column=col_index)
                cell.value = value
            else:
                print('No values detected in clip')
        
    wb.save('Excel_with_OCR_Results.xlsx') 











