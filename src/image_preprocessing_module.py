#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2

def image_preprocessing(image_path):
    '''
    # Simple image pre-processing. Reads the original colored image and converts it to grayscale
    
    Parameters
    --------------
    image_path : path/directory of original image

    Returns
    --------------
    image: pre-processed image

    '''

    ## Read Image from the given image path
    original_image  = cv2.imread(image_path)
    # Image in grayscale
    image = cv2.cvtColor(original_image, cv2.COLOR_BGR2GRAY)

    return image, original_image


