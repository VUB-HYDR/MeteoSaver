#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import matplotlib.pyplot as plt

def image_preprocessing(image_path):
    '''
    Performs image pre-processing by converting a colored image to grayscale and binarizing it by applying adaptive thresholding.

    This function reads an image from the specified path, converts it from its original colored format to grayscale, 
    and applies adaptive thresholding to binarize the image. Binarization is performed using the adaptive mean 
    thresholding method, which dynamically calculates a threshold for each small region of the image. This process 
    ensures better segmentation in varying lighting conditions.

    Parameters
    --------------
    image_path : str
        The path to the original colored image file.

    Returns
    --------------
    image_in_grayscale: numpy.ndarray
        The grayscale version of the input image.
    binarized_image: numpy.ndarray
        The pre-processed binarized image, where the foreground (text or key features) is distinguished from the background.
    original_image: numpy.ndarray
        The original colored image as read from the file.
    '''

    ## Read Image from the given image path
    original_image  = cv2.imread(image_path)
    # Convert image to grayscale
    image_in_grayscale = cv2.cvtColor(original_image, cv2.COLOR_BGR2GRAY)
    # Where: 
    
    # Apply adaptive thresholding to the grayscale image (to reduce the image to black or white pixels)
    #  The method calculates the threshold for small regions (91x91 block size) 
    # of the image using the mean of pixel intensities within that region. The threshold is then adjusted by subtracting 6. 
    # Pixels above the threshold are set to 255 (white), and those below are set to 0 (black).
    binarized_image = cv2.adaptiveThreshold(image_in_grayscale, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 91,6)
    

    ## ONLY FOR VISUALIZATION PURPOSES - UNCOMMENT THE LINES BELOW 
    ## Plotting the grayscale images
    # plt.imshow(image_in_grayscale, cmap='gray')
    # plt.title('Grayscale Image')
    # plt.axis('off')
    ## Show the plot
    # plt.show()

    # # Plotting the binarized images
    # plt.imshow(binarized_image, cmap='gray')
    # plt.title('Binarized Image')
    # plt.axis('off')
    # # Show the plot
    # plt.show()

    return image_in_grayscale, binarized_image, original_image

