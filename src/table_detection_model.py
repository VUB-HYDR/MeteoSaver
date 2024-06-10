#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import numpy as np

def detect_lines(image, kernel_size, iterations):
    # Convert to grayscale if necessary
    if len(image.shape) == 3:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    else:
        gray = image

    # Use binary thresholding
    _, img_bin = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Define a kernel for morphological operations
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, kernel_size)

    # Detect lines using morphological operations
    eroded_image = cv2.erode(img_bin, kernel, iterations=iterations)
    lines = cv2.dilate(eroded_image, kernel, iterations=iterations)

    # Find contours of the lines
    contours, _ = cv2.findContours(lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    return contours

def calculate_average_angle(contours, orientation='horizontal'):
    angles = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if orientation == 'horizontal' and w > 0:  # Avoid division by zero
            angle = np.degrees(np.arctan2(h, w))
        elif orientation == 'vertical' and h > 0:  # Avoid division by zero
            angle = np.degrees(np.arctan2(w, h))
        else:
            continue
        angles.append(angle)

    if angles:
        average_angle = np.mean(angles)
    else:
        average_angle = 0

    return average_angle

def deskew(image):

    # In this updated code:

    # The detect_horizontal_lines function detects horizontal lines in the image using morphological operations.
    # The calculate_average_angle function computes the average angle of all detected horizontal lines.
    # The deskew function rotates the image by this average angle to deskew it.
    # The table_detection function uses the deskewed image to detect the table.

    # Detect horizontal lines and calculate the average angle
    hor_contours = detect_lines(image, (np.array(image).shape[1] // 20, 1), iterations=1)
    hor_angle = calculate_average_angle(hor_contours, orientation='horizontal')

    # Rotate the image to deskew horizontally
    (h, w) = image.shape[:2]
    center = (h//2 , w//2)
    M_hor = cv2.getRotationMatrix2D(center, -hor_angle, 1.0)
    rotated_hor = cv2.warpAffine(image, M_hor, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    # # Detect vertical lines and calculate the average angle
    # ver_contours = detect_lines(rotated_hor, (1, np.array(image).shape[0] // 20), iterations=1)
    # ver_angle = calculate_average_angle(ver_contours, orientation='vertical')

    # # Rotate the image to deskew vertically
    # M_ver = cv2.getRotationMatrix2D(center, -ver_angle, 1.0)
    # rotated_ver = cv2.warpAffine(rotated_hor, M_ver, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    return rotated_hor



def table_detection(preprocessed_image, original_image):
    '''
    # Makes use of a pre-processed image (in grayscale) to detect the table from the record sheets

    Parameters
    --------------
    preprocessed_image : pre-processed image 
    original_image : original image

    Returns
    --------------
    detected_table_cells where:    
        detected_table_cells[0]: contours. Contours for the detected text in the table cells
        detected_table_cells[1]: image_with_all_bounding_boxes. Bounding boxex for each cell for which clips will be made later before optical character recognition 
        detected_table_cells[2]: table_copy. Binarized
        detected_table_cells[3]: original table image

    '''

    ## Using the pre-processing image to detect the table from the record sheets
    # Here, the threshold value for pixel intensities = 0, and the value 255 is assigned if the pixel value is above the threshold
    thresh = cv2.threshold(preprocessed_image,0,255,cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1] 
    # Perform morphological operations (like dilation and erosion) for better segmentation
    kernel = np.ones((5,5),np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # Minimum dimensions of table
    threshold_area = 4  # Minimum contour area to consider as a cell
    threshold_height = 2   # Minimum height of the cell
    threshold_width = 2    # Minimum width of the cell
    max_area = 14000 # Approx. max contour area to consider as table
    # Initialize variables for the largest contour
    largest_contour_area = 0
    largest_contour = None
    # Filter and extract individual cells, focusing on the largest contour
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        # Filter out small contours or undesired regions based on area or aspect ratio
        if cv2.contourArea(contour) > threshold_area and h > threshold_height and w > threshold_width:  # Last two conditions to filter out contours at the edges of the image
            # Find the largest contour by area
            contour_area = cv2.contourArea(contour)
            if contour_area > largest_contour_area:
                largest_contour_area = contour_area
                largest_contour = contour
    # Draw bounding box for the largest contour (if found), which here represents the table on the record sheets
    if largest_contour is not None:
        x, y, w, h = cv2.boundingRect(largest_contour)
        cv2.rectangle(preprocessed_image, (x, y), (x + w, y + h), (0, 255, 0), 2)
        table = preprocessed_image[y + 420:y + h -270 , x+200:x + w-170] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -100 is for the below the table
        table = deskew(table) # Deskew the image
        table_original_image = original_image[y:y + h, x:x + w]
        cv2.imwrite('table_original_image.jpg', table_original_image)
    else:
        table = preprocessed_image # Incase the main table is not detected as the largest contour, we just use the original image/ whole record sheet as the image with the table
        table_original_image = original_image
        cv2.imwrite('table_original_image.jpg', table_original_image)


    ## Detecting the vertical and horizontal (both dotted and bold) in the table
    # Thresholding to reduce the image to black or white pixels
    table_img_bin = cv2.adaptiveThreshold(table, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 91,6)
        
    # # Perform morphological operations to close small gaps and connect dots (dotted lines within the table)
    # kernel = np.ones((3, 3), np.uint8)
    # closing = cv2.morphologyEx(table_img_bin, cv2.MORPH_CLOSE, kernel, iterations=2)

    # Save the binary image for use later in detecting text
    cv2.imwrite('table_binarized.jpg', table_img_bin)
    #thresh,img_bin = cv2.threshold(table,100,255,cv2.THRESH_BINARY)
    img_bin = 255-table_img_bin
    # Detect the vertical lines in the image
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, np.array(table).shape[1]//50)) # The '//50' divides the length of the array (table) by 50, likely to obtain a fraction of the length for the structuring element,
    eroded_image = cv2.erode(img_bin, vertical_kernel, iterations=1)
    vertical_lines = cv2.dilate(eroded_image, vertical_kernel, iterations=5)
    # Detect the horizontal lines in the image
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(table).shape[1]//20, 1)) # The '//20' divides the width of the array (table) by 20, likely to obtain a fraction of the width for the structuring element.
    eroded_image= cv2.erode(img_bin, hor_kernel, iterations=1)
    horizontal_lines = cv2.dilate(eroded_image, hor_kernel, iterations=5)
    # Blending the imaegs with the vertical lines and the horizontal lines 
    combined_vertical_and_horizontal_lines = cv2.addWeighted(vertical_lines, 1, horizontal_lines, 1, 1)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    combined_image_dilated = cv2.dilate(combined_vertical_and_horizontal_lines, kernel, iterations=5)
    # Remove the lines from the image (table)
    image_without_lines = cv2.subtract(img_bin, combined_image_dilated)
    # Remove smaller 'still-visible' lines through noise removal
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    image_without_lines_noise_removed = cv2.erode(image_without_lines, kernel, iterations=1)
    image_without_lines_noise_removed = cv2.dilate(image_without_lines_noise_removed, kernel, iterations=1)
    # Convert words into blobs using dilation
    #kernel_to_remove_gaps_between_words = np.ones((5, 5), np.uint8)  # Larger kernel to bridge gaps better
    kernel_to_remove_gaps_between_words = np.array([
            [1,1,1,1,1, 1],
            [1,1,1,1,1, 1]])
    image_with_word_blobs = cv2.dilate(image_without_lines_noise_removed, kernel_to_remove_gaps_between_words, iterations=5) # was 5 iterations previously
    
    # # Apply morphological closing to close gaps between letters within words
    # closing_kernel = np.ones((5, 5), np.uint8)
    # image_with_word_blobs = cv2.morphologyEx(image_with_word_blobs, cv2.MORPH_CLOSE, closing_kernel)

    simple_kernel = np.ones((3,3), np.uint8)
    image_with_word_blobs = cv2.dilate(image_with_word_blobs, simple_kernel, iterations=1)
    # Detecting the dotted lines using horizontal line detection and erosion. ### ADDITIONAL STEP: This is because the original images ahve dotted horizontal lines which cvan still be detected after the first removal of main (undotted) horizontal lines
    hor_kernel_2 = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(image_with_word_blobs).shape[1]//20, 1))
    image_3 = cv2.erode(image_with_word_blobs, hor_kernel_2, iterations=1)
    horizontal_lines_2 = cv2.dilate(image_3, hor_kernel_2, iterations=1)
    # Removing the dotted liens by substracting them from the original image 
    image_without_lines_2 = cv2.subtract(image_with_word_blobs, horizontal_lines_2)



    ## Using contours in order to detect text in the table after removing the vertical and horizontal lines
    # Assuming 'table' is your input image in BGR format
    result = cv2.findContours(image_without_lines_2, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contours = result[0]
    # Original image of table in binarizesd format
    image_with_all_bounding_boxes = cv2.imread('table_binarized.jpg')
    table_binarized = image_with_all_bounding_boxes.copy()
    

    detected_table_cells = [contours, image_with_all_bounding_boxes, table_binarized, table_original_image]
    
    return detected_table_cells

