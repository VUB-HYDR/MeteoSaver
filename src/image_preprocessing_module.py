#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2

def image_preprocessing(image_path):
    '''
    Performs basic image pre-processing by converting a colored image to grayscale.

    This function reads an image from the specified path and converts it from its original colored format to grayscale. 
    This is a common preprocessing step in image analysis and computer vision tasks, where color information may not be necessary.

    Parameters
    --------------
    image_path : str
        The path to the original colored image file.

    Returns
    --------------
    image: numpy.ndarray
        The pre-processed grayscale image.
    original_image: numpy.ndarray
        The original colored image as read from the file.
    '''

    ## Read Image from the given image path
    original_image  = cv2.imread(image_path)
    # Image in grayscale
    image = cv2.cvtColor(original_image, cv2.COLOR_BGR2GRAY)

    return image, original_image


