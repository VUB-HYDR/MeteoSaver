#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import numpy as np
import matplotlib.pyplot as plt

def detect_lines(image, kernel_size, iterations):
    '''
    Detects lines in an image using morphological operations and returns the contours of the detected lines.

    This function processes an input image to detect lines by applying binary thresholding followed by a series of morphological operations (erosion and dilation). It first converts the image to grayscale if necessary, then uses a rectangular structuring element to enhance line structures in the image. The resulting lines are detected by finding contours on the processed image.

    Parameters
    --------------
    image : 
        The input image in which lines are to be detected. The image can be in grayscale or BGR format; if in BGR format, it will be converted to grayscale.
    
    kernel_size : tuple of int
        The size of the structuring element used for the morphological operations. This tuple determines the dimensions of the rectangular kernel (width, height) that will be used for erosion and dilation.
    
    iterations : int
        The number of times the morphological operations (erosion and dilation) will be applied. Increasing the number of iterations can help to connect broken lines or separate closely spaced lines.

    Returns
    --------------
    contours : list of numpy.ndarray
        A list of contours representing the detected lines in the image. Each contour is an array of points that outline a detected line.
    '''


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
    '''
    Calculates the average angle of contours relative to the specified orientation (horizontal or vertical).

    This function computes the angles of a set of contours based on their bounding rectangles. Depending on the specified orientation, it calculates the angle between the width and height of each contour's bounding box. The function then returns the average angle of all contours, which can provide insight into the overall alignment or skewness of the detected shapes.

    Parameters
    --------------
    contours : list
        A list of contours, where each contour is an array of points representing a detected shape in the image.
    
    orientation : str, optional
        The reference orientation for calculating angles. Accepts 'horizontal' or 'vertical'.
        - 'horizontal': The angle is calculated relative to the horizontal axis (based on width and height).
        - 'vertical': The angle is calculated relative to the vertical axis (based on height and width).
        The default value is 'horizontal'.

    Returns
    --------------
    average_angle : float
        The average angle of the contours relative to the specified orientation, measured in degrees.
        If no valid angles are found, the function returns 0.
    '''

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
    '''
    Deskews an image by detecting and correcting its skew based on the orientation of detected horizontal lines.

    This function corrects the skew of an input image by first detecting horizontal lines within the image using morphological operations. It calculates the average angle of these detected lines and rotates the image by this angle to align the horizontal lines correctly, effectively deskewing the image. The result is an image where the content is horizontally aligned, which is particularly useful for preprocessing before further analysis or OCR (Optical Character Recognition).

    Parameters
    --------------
    image : 
        The input image that needs to be deskewed. This image can be in grayscale or color format.

    Returns
    --------------
    rotated_hor : 
        The deskewed image after rotation to correct horizontal alignment. The output image is rotated by the calculated average angle of the detected horizontal lines.
    '''


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



def table_detection(preprocessed_image, original_image, clip_up, clip_down, clip_left, clip_right):
    '''
    Detects and extracts a table from a pre-processed grayscale image, returning the contours and relevant images for further processing.

    This function processes a pre-processed grayscale image to detect and extract a table, which is then used for optical character recognition (OCR) or other analytical purposes. The function applies various image processing techniques, including thresholding, morphological operations, and contour detection, to isolate the table from the background. 
    It also removes lines from the table to isolate the text within the cells. The resulting data includes the contours of detected table cells, images with bounding boxes around each cell, and the original and binarized versions of the detected table.

    Parameters
    --------------
    preprocessed_image : 
        The pre-processed grayscale image where the table detection will be performed. This image is alreasy in grayscale
    
    original_image : 
        The original image corresponding to the preprocessed image. This is used for extracting the original table image without any preprocessing artifacts.
    
    clip_up : int
        The number of pixels to clip from the top of the DETECTED TABLE for further processing.
    
    clip_down : int
        The number of pixels to clip from the bottom of the DETECTED TABLE for further processing.
    
    clip_left : int
        The number of pixels to clip from the left side of the DETECTED TABLE for further processing.
    
    clip_right : int
        The number of pixels to clip from the right side of the DETECTED TABLE for further processing.

    Here the clip_up, clip_down, clip_left, and clip_right ensure clipping of the HEADERS and ROW LABELS (Date & Pentad no. in our case) from the entire detected table (table detected using ML). Adjust this to your case study. Incase you would like to maintain the full table, set clip_up, clip_down, clip_left, clip_right = 0

    Returns
    --------------
    detected_table_cells : list
        A list containing:
        - detected_table_cells[0]: contours representing the detected text in the table cells.
        - detected_table_cells[1]: image_with_all_bounding_boxes containing the bounding boxes drawn around each detected cell.
        - detected_table_cells[2]: table_copy, which is the binarized version of the detected table.
        - detected_table_cells[3]: table_original_image, which is the clipped original image of the detected table.
        - detected_table_cells[4]: full_detected_table_with_labels, which includes the full table with labels before any clipping.
    '''


    # Here, we employ ML algorithms from Open Source Computer Vision (OpenCV) following methodologies similar to those described in https://livefiredev.com/how-to-extract-table-from-image-in-python-opencv-ocr/ [GitHub repository: \url{https://github.com/livefiredev/ocr-extract-table-from-image-python}, (last access: 19 July 2024)], but further customizing them for our case study.

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
        full_detected_table_with_labels = preprocessed_image[y:y + h, x:x + w] 
        full_detected_table_with_labels = cv2.adaptiveThreshold(full_detected_table_with_labels, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 91,6) # Thresholding to reduce the image to black or white pixels
        table = preprocessed_image[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
        #table = deskew(table) # Deskew the image, # Optional: Incase some of your images are skewed.
        table_original_image = original_image[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
        # cv2.imwrite('table_original_image.jpg', table_original_image)
    else:
        table = preprocessed_image # Incase the main table is not detected as the largest contour, we just use the original image/ whole record sheet as the image with the table
        table_original_image = original_image
        full_detected_table_with_labels = cv2.adaptiveThreshold(table, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 91,6) # Thresholding to reduce the image to black or white pixels
        # cv2.imwrite('table_original_image.jpg', table_original_image)


    ## Detecting the vertical and horizontal (both dotted and bold) in the table
    # Thresholding to reduce the image to black or white pixels
    table_img_bin = cv2.adaptiveThreshold(table, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 91,6)

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
    

    detected_table_cells = [contours, image_with_all_bounding_boxes, table_binarized, table_original_image, full_detected_table_with_labels]
    
    # Plots only for visualization purposes. Uncomment the lines below to show the different steps

    # plt.imshow(image_without_lines_noise_removed, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.show() 

    # plt.imshow(image_without_lines_2, cmap = 'gray') # figure showing text blobs on the detected table image with horizintal and vertical lines removed.
    # plt.show()

    # plt.imshow(detected_table_cells[4], cmap = 'gray') # unclipped detected table
    # plt.show()

    # plt.imshow(detected_table_cells[1]) # clipped detected table
    # plt.show()

    return detected_table_cells

