#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
# from paddleocr import PaddleOCR,draw_ocr
from PIL import Image, ImageDraw, ImageFont
import pytesseract
import easyocr
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
import math
import random
from scipy.stats import trim_mean

def organize_contours_midpoint(contours, max_rows):
    '''
    # Organizes the bounding boxes, here termed as contours, in rows by their center co-ordinates using the KMeans clustering

    Parameters
    --------------
    contours: list of bounding boxes (with their x, y, w, h coordinates)
        List of contours for the detected text in the table cells with coordinates

    max_rows: int
        Maximum rows, adjusted based on your table's expected structure

    Returns
    -------------- 
    rows: Bounding boxes organised in rows using Kmeans clustering

    '''

    midpoints = [(cv2.boundingRect(contour)[1] + cv2.boundingRect(contour)[3] // 2) for contour in contours]
    if len(midpoints) == 0:
        return []
    kmeans = KMeans(n_clusters=min(max_rows, len(midpoints)), random_state=0)
    kmeans.fit(np.array(midpoints).reshape(-1, 1))
    labels = kmeans.labels_
    rows = [[] for _ in range(max_rows)]
    for label, contour in zip(labels, contours):
        rows[label].append(contour)
    for i in range(len(rows)):
        rows[i] = sorted(rows[i], key=lambda c: cv2.boundingRect(c)[0])
    return rows


def organize_contours_top(contours, max_rows):
    '''
    # Organizes the bounding boxes, here termed as contours, in rows by their top co-ordinates using the KMeans clustering

    Parameters
    --------------
    contours: list of bounding boxes (with their x, y, w, h coordinates)
        List of contours for the detected text in the table cells with coordinates

    max_rows: int
        Maximum rows, adjusted based on your table's expected structure

    Returns
    -------------- 
    rows: Bounding boxes organised in rows using Kmeans clustering

    '''

    # Top for vertical clustering
    top = [cv2.boundingRect(contour)[1] for contour in contours]
    if len(top) == 0:
        return []

    kmeans = KMeans(n_clusters=min(max_rows, len(top)), random_state=0)
    kmeans.fit(np.array(top).reshape(-1, 1))
    labels = kmeans.labels_

    rows = [[] for _ in range(max_rows)]
    for label, contour in zip(labels, contours):
        rows[label].append(contour)

    for i in range(len(rows)):
        rows[i] = sorted(rows[i], key=lambda c: cv2.boundingRect(c)[0])

    return rows


def organize_contours_fraction(contours, max_rows, fraction=0.33):
    '''
    Organizes the bounding boxes (contours) in rows by a fractional point between the top and midpoint using KMeans clustering.

    Parameters
    --------------
    contours: list of bounding boxes (with their x, y, w, h coordinates)
        List of contours for the detected text in the table cells with coordinates.

    max_rows: int
        Maximum rows, adjusted based on your table's expected structure.

    fraction: float, optional (default=0.33)
        The fractional point of the height from the top for vertical clustering. A value between 0 (top) and 1 (bottom).
    
    Returns
    --------------
    rows: Bounding boxes organized in rows using KMeans clustering.
    '''

    # Calculate the point at the specified fraction between the top and bottom
    fraction_points = [(cv2.boundingRect(contour)[1] + int(cv2.boundingRect(contour)[3] * fraction)) for contour in contours]
    
    if len(fraction_points) == 0:
        return []

    # Apply KMeans clustering based on the fractional points
    kmeans = KMeans(n_clusters=min(max_rows, len(fraction_points)), random_state=0)
    kmeans.fit(np.array(fraction_points).reshape(-1, 1))
    labels = kmeans.labels_

    rows = [[] for _ in range(max_rows)]
    for label, contour in zip(labels, contours):
        rows[label].append(contour)

    # Sort each row by the x-coordinate for left-to-right ordering within rows
    for i in range(len(rows)):
        rows[i] = sorted(rows[i], key=lambda c: cv2.boundingRect(c)[0])

    return rows

def calculate_cell_reference(x, row_index, max_columns, table_width):
    '''
    Calculates the Excel cell reference (e.g., 'B3') based on the horizontal position of a detected cell's x-coordinate within a table.

    This function determines the appropriate Excel cell reference by converting the horizontal position of a cell's x-coordinate into a column index. The column index is combined with the given row index to generate the full cell reference in the format used by Excel (e.g., 'A1', 'B2'). The column index is adjusted to ensure it remains within valid boundaries (1 to `max_columns`).

    Parameters
    --------------
    x : float or int
        The x-coordinate of the cell (in pixels), representing its horizontal position within the table.
    
    row_index : int
        The row number in the table to which the cell belongs. This is used directly in the final cell reference.
    
    max_columns : int
        The maximum number of columns in the table. This ensures that the calculated column index does not exceed the available columns.
    
    table_width : int or float
        The total width of the table (in pixels). This is used to determine the relative position of `x` within the table and calculate the corresponding column index.

    Returns
    --------------
    cell_reference : str
        The Excel-style cell reference (e.g., 'B3', 'C7') corresponding to the detected cell's position within the table.
    '''

    column = math.floor(x / table_width * max_columns) + 1

    # Ensure the column index is within valid ranges
    if column < 1:
        column = 1
    elif column > max_columns:
        column = max_columns

    return f'{openpyxl.utils.get_column_letter(column)}{row_index}'




def generate_random_colors(n):
    '''
    Generates a list of `n` random distinct colors in BGR format.

    This function creates a list of distinct colors by generating random values for hue, saturation, and value (HSV) and then converting them to the BGR color space, which is commonly used in OpenCV. The generated colors are vivid and bright, ensuring they are visually distinct when used in applications such as object detection, image segmentation, or visual annotations.

    Parameters
    --------------
    n : int
        The number of distinct colors to generate.

    Returns
    --------------
    colors : list of tuples
        A list containing `n` tuples, each representing a color in BGR format. The colors are designed to be vivid and distinct for clear visualization.
    '''


    colors = []
    for i in range(n):
        # Generate random values for hue, saturation, and value
        hue = random.randint(0, 179)  # Hue range in OpenCV is [0, 179]
        saturation = random.randint(100, 255)  # To ensure the colors are vivid
        value = random.randint(100, 255)  # To ensure the colors are bright

        # Convert the random HSV color to BGR
        color = cv2.cvtColor(np.uint8([[[hue, saturation, value]]]), cv2.COLOR_HSV2BGR)[0][0].tolist()
        colors.append((int(color[0]), int(color[1]), int(color[2])))
    return colors




def draw_row_markers_and_boxes(image, rows, colors):
    '''
    Draws bounding boxes around contours in each row and adds numbered markers to each row with distinct colors.

    This function takes an image and a list of rows, where each row is a list of contours representing table cells or regions of interest. It draws a bounding box around each contour in the row and places a numbered marker to the left of the row. Each row is highlighted with a distinct color, cycling through the provided color list if necessary.

    Parameters
    --------------
    image : 
        The image on which the bounding boxes and row markers will be drawn. This image is modified in place.
    
    rows : list of lists
        A list of rows, where each row is a list of contours (represented as arrays of points) that are part of the same row in a table or grid.
    
    colors : list of tuples
        A list of BGR color tuples used to color the bounding boxes and markers for each row. The function cycles through this list if there are more rows than colors.

    Returns
    --------------
    None
        The function modifies the input image in place and does not return any value. The image will have bounding boxes drawn around the contours and numbered markers for each row.
    '''


    font = cv2.FONT_HERSHEY_SIMPLEX
    font_scale = 0.5
    thickness = 1

    for idx, row in enumerate(rows):
        color = colors[idx % len(colors)]  # Cycle through colors if more rows than colors
        y_coords = [cv2.boundingRect(c)[1] + cv2.boundingRect(c)[3] // 2 for c in row]
        if y_coords:
            y_position = int(np.median(y_coords))
            x_position = 10  # arbitrary x position to place the marker
            cv2.putText(image, str(idx + 1), (x_position, y_position), font, font_scale, color, thickness)
        
        # Draw bounding boxes for each contour in the row
        for contour in row:
            x, y, w, h = cv2.boundingRect(contour)
            cv2.rectangle(image, (x, y), (x + w, y + h), color, 2)




def calculate_trimmed_mean(values, proportion_to_cut=0.2):
    '''
    Calculates the trimmed mean of a list of values, excluding a specified proportion of the smallest and largest values.

    This function computes the trimmed mean of a given list of numerical values. The trimmed mean is a measure of central tendency that removes a certain proportion of the lowest and highest values before calculating the mean. This helps to reduce the effect of outliers on the mean.

    Parameters
    --------------
    values : list
        A list of numerical values for which the trimmed mean is to be calculated.
    
    proportion_to_cut : float, optional
        The proportion of values to remove from each end of the sorted list before calculating the mean. 
        For example, a proportion of 0.2 means removing the lowest 20% and the highest 20% of the values. 
        The default value is 0.2.

    Returns
    --------------
    float
        The trimmed mean of the input values, after excluding the specified proportion of extreme values.
    '''

    return trim_mean(values, proportion_to_cut)


def merge_excel_files(file1, file2, output_file, start_row, end_row):
    '''
    Merges two Excel files for verification purposes: one organized by the mid-point coordinates of bounding boxes and the other by the top coordinates.

    This function merges two preprocessed Excel files that contain transcribed data organized differently (one by mid-point and the other by top coordinates of bounding boxes). 
    The merged output file allows for cross-checking to ensure that cells are correctly placed in their respective rows.

    Parameters
    --------------
    file1: str
        The path to the Excel file containing transcribed data organized in rows using the top coordinates of the bounding boxes (contours).
    file2: str
        The path to the Excel file containing transcribed data organized in rows using the mid-point coordinates of the bounding boxes (contours).
    output_file: str
        The path where the merged Excel file will be saved.
    start_row: int
        The starting row number from which to begin the merge.
    end_row: int
        The ending row number up to which the merge should be conducted.

    Returns
    --------------
    None
        The function creates and saves a merged Excel file at the specified `output_file` location. This file combines the data from `file1` and `file2` for further verification.
    '''


    # Load the Excel files into DataFrames without headers
    df1 = pd.read_excel(file1, header=None)
    df2 = pd.read_excel(file2, header=None)


    # If the indices are not simple integers or do not align with Excel rows as expected,
    # you might need to reset them or adjust how the Excel file is being read (e.g., `index_col=None`)
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)

    # Convert start_row and end_row to zero-based index for Python
    start_idx = start_row - 1  # Convert 1-based index to 0-based
    end_idx = end_row - 1  # Convert 1-based index to 0-based

    # Slice to only include the range from start_idx to end_idx
    df1 = df1.iloc[start_idx:end_idx + 1]
    df2 = df2.iloc[start_idx:end_idx + 1]

    # Initialize a new DataFrame to hold merged results
    merged_df = pd.DataFrame(index=df1.index, columns=df1.columns)

    # Iterate over rows by index (assuming the indices are aligned)
    for idx in df1.index:
        for col in df1.columns:
            val1 = df1.at[idx, col]
            val2 = df2.at[idx, col]
            # Simple merge logic: prefer non-empty values from df1, then df2
            if pd.notna(val1):
                merged_df.at[idx, col] = val1
            else:
                merged_df.at[idx, col] = val2


    # Create a new workbook and select the active worksheet
    new_workbook = openpyxl.Workbook()
    new_worksheet = new_workbook.active

    # Append the merged DataFrame to the new worksheet
    for r_idx, row in enumerate(dataframe_to_rows(merged_df, index=False, header=False), start=1):
        for c_idx, value in enumerate(row, start=1):
            new_worksheet.cell(row=r_idx, column=c_idx, value=value)

    # Merge cells for multi-column headers
    new_worksheet.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1) #No de la pentade
    new_worksheet.merge_cells(start_row=1, start_column=2, end_row=3, end_column=2) #Date
    new_worksheet.merge_cells(start_row=1, start_column=3, end_row=3, end_column=3) #Bellani
    new_worksheet.merge_cells(start_row=1, start_column=4, end_row=1, end_column=8) #Températures extrêmes
    new_worksheet.merge_cells(start_row=1, start_column=9, end_row=1, end_column=10) #Evaportation
    new_worksheet.merge_cells(start_row=1, start_column=11, end_row=3, end_column=11) #Pluies
    new_worksheet.merge_cells(start_row=1, start_column=12, end_row=1, end_column=16) #Température et Humidité de l'air à 6 heures
    new_worksheet.merge_cells(start_row=1, start_column=17, end_row=1, end_column=21) #Température et Humidité de l'air à 15 heures
    new_worksheet.merge_cells(start_row=1, start_column=22, end_row=1, end_column=26) #Température et Humidité de l'air à 18 heures
    new_worksheet.merge_cells(start_row=1, start_column=27, end_row=3, end_column=27) #Date
    # subheaders
    new_worksheet.merge_cells(start_row=2, start_column=4, end_row=2, end_column=7) #Abri
    new_worksheet.merge_cells(start_row=2, start_column=9, end_row=2, end_column=10) #Piche
    new_worksheet.merge_cells(start_row=2, start_column=12, end_row=2, end_column=16) #(Psychromètre a aspiration)
    new_worksheet.merge_cells(start_row=2, start_column=17, end_row=2, end_column=21) #(Psychromètre a aspiration)
    new_worksheet.merge_cells(start_row=2, start_column=22, end_row=2, end_column=26) #(Psychromètre a aspiration)

    # Set up border styles for excel output
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'))

    # Loop through cells to apply borders
    for row in new_worksheet.iter_rows(min_row=1, max_row=new_worksheet.max_row, min_col=1, max_col=new_worksheet.max_column):
        for cell in row:
            cell.border = thin_border
    new_workbook.save(output_file)
    
    # Iterate through all cells and set the alignment
    for row in new_worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Read headers from the first row of one of the files
    workbook = openpyxl.load_workbook(file1)
    copy_file1 = workbook.active
    headers = [cell.value for cell in copy_file1[1]]  
    for row in new_worksheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=new_worksheet.max_column):
        for col_num, header in enumerate(headers, start=1):
            cell = new_worksheet.cell(row=1, column=col_num, value=header)
            if header == "No de la pentade" or header == "Date" or header == "Bellani (gr. Cal/cm2) 6-6h" or header == "Pluies en mm. 6-6h":
                cell.alignment = Alignment(textRotation=90)

    # Save the workbook
    new_workbook.save(output_file)



def transcription(detected_table_cells, ocr_model, tesseract_path, transient_transcription_output_dir, pre_QA_QC_transcribed_hydroclimate_data_dir_station, station, month_filename, no_of_rows, no_of_columns, no_of_rows_including_headers):
    '''
    Performs OCR (Optical Character Recognition) on detected table cells from a pre-processed grayscale image to extract text and store the results in an Excel workbook.

    This function processes a detected table within an image, using contours to identify individual cells. It clips each cell and applies the specified OCR/HTR (Handwritten Text Recognition) model to transcribe the text within each cell. The results are then saved in an Excel workbook.

    Parameters
    --------------
    detected_table_cells : list
        detected_table_cells[0]: contours. Contours for the detected text in the table cells.
        detected_table_cells[1]: image_with_all_bounding_boxes. Image with bounding boxes drawn around each detected table cell.
        detected_table_cells[2]: table_copy. A copy of the processed table image used for further operations.
        detected_table_cells[3]: table_original_image. The original image of the table before any processing.
    
    ocr_model : string
        The OCR or HTR model used to recognize and transcribe text from the detected table cells. This can be any model that supports text recognition, such as Tesseract, EasyOCR, or a custom deep learning model.
    
    no_of_rows : int
        The number of rows in the detected table. This parameter helps in structuring the data correctly in the output workbook.
    
    no_of_columns : int
        The number of columns in the table, used to structure the OCR results.
    
    min_cell_width_threshold : int
        The minimum width of a cell (in pixels) that should be considered for OCR. Cells smaller than this threshold are ignored.
    
    max_cell_width_threshold : int
        The maximum width of a cell (in pixels) that should be considered for OCR. Cells wider than this threshold are ignored.
    
    min_cell_height_threshold : int
        The minimum height of a cell (in pixels) that should be considered for OCR. Cells smaller than this threshold are ignored.
    
    max_cell_height_threshold : int
        The maximum height of a cell (in pixels) that should be considered for OCR. Cells taller than this threshold are ignored.

    Returns
    -------------- 
    wb : openpyxl.Workbook
        An Excel workbook object where the OCR results are stored. Each cell's transcribed text is placed in the corresponding cell of the Excel sheet based on its detected position in the image.
    '''
    
    
    if ocr_model == 'Tesseract-OCR':
        ## Lauching Tesseract-OCR
        pytesseract.pytesseract.tesseract_cmd = tesseract_path ## Here input the PATH to the Tesseract executable on your computer. See more information here: https://pypi.org/project/pytesseract/
    # if ocr_model == 'PaddleOCR':
    #     ## Lauching PaddleOCR, which would be used by downloading necessary files as shown below
    #     paddle_ocr = PaddleOCR(use_angle_cls=True, lang = 'en', use_gpu=False) ## Run only once to download all required files
    if ocr_model == 'EasyOCR':
        ## Lauching EasyOCR
        easyocr_reader = easyocr.Reader(['en']) # this needs to run only once to load the model into memory

    easyocr_reader = easyocr.Reader(['en'])

    # Contours of bounding boxes detected in cell recognition
    new_contours = detected_table_cells[0]

    image_with_all_bounding_boxes = detected_table_cells[1]
    table_copy = detected_table_cells[2]

    # Get the dimensions of the loaded image. Here, particulary the image/table width is very important for the column placement of cells/bounding boxes
    image_height, image_width, image_channels = image_with_all_bounding_boxes.shape

    # Image onto which the regions of interest (ROIs) i.e the cells, will be drawn for illustration purposes
    ROIs_image = table_copy.copy()
    
    results = []

    # Here we use two methods to arrange the boundung boxes (as a double check): (1) Using the middle coordinates of the bounding boxes , and (2) Using the top coordinated of the boudning boxes.
    organize_methods = {  
        'Midpoint': organize_contours_midpoint,
        'Top': organize_contours_top
    }

    for method_name, organize_method in organize_methods.items():

        ## Create an Excel workbook and add a worksheet where the transcribed text will be saved
        wb = Workbook()
        ws = wb.active
        ws.title = 'OCR_Results'

        # Organize the contours
        organized_rows = organize_method(new_contours, no_of_rows)
        # Sorting the cell (bounding box) rows from first to last using trimmed mean of coordinates
        sorted_rows = sorted(organized_rows, key=lambda row: calculate_trimmed_mean([cv2.boundingRect(c)[1] + cv2.boundingRect(c)[3] // 2 for c in row]))

    
        ### FOR ILLUSTRATION PURPOSES: Uncomment the following lines if you want to see an example of how the sorting function works. This is for illustration purposes only and not necessary for the main functionality of the script. 
        
        # # Create a copy of the image for visualization
        # image_before_sorting = table_copy.copy()
        # image_after_sorting = table_copy.copy()

        # # Generate colors for 43 rows. For sorted row visualitation purposes
        # colors = generate_random_colors(no_of_rows)  # Random colours for the maximum number of rows, such that each row has its own color for easy identification

        # # Draw markers on the image before sorting
        # draw_row_markers_and_boxes(image_before_sorting, organized_rows, colors)  # Green color for original order

        # # Draw markers on the image after sorting
        # draw_row_markers_and_boxes(image_after_sorting, sorted_rows, colors)  # Red color for sorted order

        # # Save or display the images for inspection
        # cv2.imwrite(f'before_sorting_{method_name}.png', image_before_sorting)  # or use cv2.imshow and cv2.waitKey for immediate display
        # plt.imshow('before_sorting.png')
        # plt.show()

        # cv2.imwrite(f'after_sorting_{method_name}.png', image_after_sorting)  # or use cv2.imshow and cv2.waitKey for immediate display
        # plt.imshow('after_sorting.png')
        # plt.show()
        
    

        # Sort boxes within each column of each row by y-coordinate
        for row in sorted_rows:
            row.sort(key=lambda c: cv2.boundingRect(c)[1])

        for row_index, row in enumerate(sorted_rows, start=1):
            
            for contour in row:

                x, y, w, h = cv2.boundingRect(contour)              

                # Calculate a factor to modify the bounding box area (e.g., in this case 5% of width and 25% of height)
                factor_width = 0.05  # Modify this factor as needed
                increase_factor_height = 0.25 # Modify this factor as needed
            
                x += int(w * factor_width)  # Increase width
                y -= int(h * increase_factor_height)  # Increase height
                w -= int(1 * w * factor_width)  # Decrease width a little to avoid vertical lines that may be transcribed as the number 1 yet they aren't a number
                h += int(2 * h * increase_factor_height)  # Increase height
                
                ### FOR ILLUSTRATION PURPOSES: This line below is about drawing a rectangle on the image with the shape of the bounding box. Its not needed for the OCR. This is only for debugging purposes.
                # image_with_all_bounding_boxes = cv2.rectangle(image_with_all_bounding_boxes, (x, y), (x + w, y + h), (0, 255, 0), 5)

                # Draw the adjusted ROI on the output image
                cv2.rectangle(ROIs_image, (x, y), (x + w, y + h), (0, 255, 0), 5)  # (0, 255, 0) represent a green color for ROI, and 5 is the thicnkess of the ROI bounbdary boxes
                
                
                # OCR
                # Crop each cell using the bounding rectangle coordinates
                ROI = table_copy[y:y+h, x:x+w] # Here, the Region Of Interest (ROI) represent the cells (boundary boxes) clipped out from the table image as a prerequiste for text recogniton by the Optical Character Recognition/Handwritten Text Recognition (OCR/HTR) model
                
                if ROI.size != 0:  # Check if the height and width are greater than zero. This is to prevent invalid ROIs
                    
                    # Save the detected text image/ROI
                    save_dir =  os.path.join(transient_transcription_output_dir, station)
                    os.makedirs(save_dir, exist_ok=True)  # Ensure the directory exists
                    save_path_detected_text = os.path.join(save_dir, 'detected.png')
                    cv2.imwrite(save_path_detected_text, ROI)
                    if ocr_model == 'Tesseract-OCR':
                    # Using Tesseract-OCR
                        ocr_result = pytesseract.image_to_string(save_path_detected_text, lang='cobecore-V6', config='--psm 7 -c tessedit_char_whitelist=0123456789') # Just added -c tessedit_char_whitelist=0123456789 to really limit the text type/values detected

                        # Here's a brief explanation of some Page Segmentation Modes (PSMs) available in Tesseract:
                        # 0: Orientation and script detection (OSD) only.
                        # 1: Automatic page segmentation with OSD.
                        # 2: Automatic page segmentation, but no OSD, or OCR.
                        # 3: Fully automatic page segmentation, but no OSD. (Default)
                        # 4: Assume a single column of text of variable sizes.
                        # 5: Assume a single uniform block of vertically aligned text.
                        # 6: Assume a single uniform block of text.
                        # 7: Treat the image as a single text line.
                        # 8: Treat the image as a single word.
                        # 9: Treat the image as a single word in a circle.
                        # 10: Treat the image as a single character.
                        # 11: Sparse text. Find as much text as possible in no particular order.
                        # 12: Sparse text with OSD.
                        # 13: Raw line. Treat the image as a single text line, bypassing hacks that are Tesseract-specific.

                    # Uncomment the following lines if you'd like to make use of PaddleOCR
                    # if ocr_model == 'PaddleOCR':
                    #     ## Using PaddleOCR
                    #     ocr_result = paddle_ocr.ocr('detected.png', cls = True)

                    if ocr_model == 'EasyOCR':
                    # Using EasyOCR
                        ocr_result = easyocr_reader.readtext(save_path_detected_text, detail = 0, allowlist='0123456789')
                        # In EasyOCR, the detail parameter specifies the level of detail in the output. 
                                # When using the readtext method, the detail parameter can be set to different values to control what kind of output you get. 
                                # Specifically:
                                # detail=1: The output will be a list of tuples, where each tuple contains detailed information about the detected text, 
                                # including the bounding box coordinates, the text string, and the confidence score. Example: [(bbox, text, confidence), ...].

                                # detail=0: The output will be a list of strings, where each string is the detected text without any additional details. 
                                # This is a simpler output format that only provides the recognized text. Example: ["text1", "text2", ...].
                        if isinstance(ocr_result, list): # This is because EasyOCR's results are returned as a list
                                ocr_result = ''.join(ocr_result)  # Convert list to a string
                    
                    
                    # Using OCR for handwritten text recognition
                    if ocr_result is not None:
                        
                        if not ocr_result.strip(): # Check if the result is empty or only whitespace. This could be due to the selected OCR (in this case: Tesseract-OCR) not being able to recognize the text in the ROI.
                            # For this reason, we can try another OCR, say for example Easy OCR, to try to recognize the text in this ROI
                            ocr_result = easyocr_reader.readtext(save_path_detected_text, detail = 0, allowlist='0123456789')
                            if isinstance(ocr_result, list): # This is because EasyOCR's results are returned as a list
                                ocr_result = ''.join(ocr_result)  # Convert list to a string


                                               
                        # Attain the Ms Excel Template cell coordinates
                        # Determine the cell reference using the x coodrinate of the bounding box, row index, maximum column number, and image/table width
                        cell_ref = calculate_cell_reference(x, row_index, max_columns=24, table_width=image_width) # e.g., A1, B5, etc.
                        
                        # Additional check: Incase the cell (cell_ref) is already occupied with transcribed text. We then opt for the cell below within the same column
                        column_letter = openpyxl.utils.get_column_letter(math.floor(x / image_width * no_of_columns) + 1)
                        initial_row_index = row_index  # Store the initial row index
                        # Check if the cell is already occupied
                        if ws[cell_ref].value is not None:
                            row_index += 1
                            cell_ref = f'{column_letter}{row_index}'
                        
                        # Place the OCR/HTR recognized text in its respective Ms Excel cell 
                        ws[cell_ref].value = ocr_result   

                        # Restore the row index to the initial value
                        row_index = initial_row_index

                        # Set up border styles for excel output
                        thin_border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin'))

                        # Loop through cells to apply borders
                        for row in ws.iter_rows(min_row=1, max_row=no_of_rows, min_col=1, max_col=no_of_columns):
                            for cell in row:
                                cell.border = thin_border
                    

                    else:
                        print('No values detected in clip')
                else:
                    print('ROI is empty or invalid')


        ### SPECIAL ADDITION TO CODE / CUSTOMIZATION: The lines below are particular to our tables. Here we replace the headers and row labels as in the original table format. This will vary from your setup and this should be customized to your particular sheets.
        # Insert two columns on the left side of the excel sheet
        ws.insert_cols(1, 2)
        
        # Insert a new row at the top for headers
        ws.insert_rows(1, amount = 3)

        # Define your headers (adjust as needed)
        headers = ["No de la pentade", "Date", "Bellani (gr. Cal/cm2) 6-6h", "Températures extrêmes", "", "", "", "", "Evaportation en cm3 6 - 6h", "", "Pluies en mm. 6-6h", "Température et Humidité de l'air à 6 heures", "", "", "", "", "Température et Humidité de l'air à 15 heures",  "", "", "", "", "Température et Humidité de l'air à 18 heures",  "", "", "", "", "Date"]
        sub_headers_1 = ["", "", "", "Abri", "", "", "", "", "Piche", "", "", "(Psychromètre a aspiration)", "", "", "", "", "(Psychromètre a aspiration)", "", "", "", "", "(Psychromètre a aspiration)", "", "", "", ""]
        sub_headers_2 =["", "", "", "Max.", "Min.", "(M+m)/2", "Ampl.", "Min. gazon", "Abri.", "Ext.", "", "T", "T'a", "e.", "U", "∆e", "T", "T'a", "e.", "U", "∆e","T", "T'a", "e.", "U", "∆e", ""]

        # Add the headers to the first row
        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_num, value=header)
            if header == "No de la pentade" or header == "Date" or header == "Bellani (gr. Cal/cm2) 6-6h" or header == "Pluies en mm. 6-6h":
                cell.alignment = Alignment(textRotation=90)

        # Add the first row of sub-headers to the second row
        for col_num, sub_header in enumerate(sub_headers_1, start=1):
            ws.cell(row=2, column=col_num, value=sub_header)

        # Add the second row of sub-headers to the third row
        for col_num, sub_header in enumerate(sub_headers_2, start=1):
            ws.cell(row=3, column=col_num, value=sub_header)
        
        # Merge cells for multi-column headers
        ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1) #No de la pentade
        ws.merge_cells(start_row=1, start_column=2, end_row=3, end_column=2) #Date
        ws.merge_cells(start_row=1, start_column=3, end_row=3, end_column=3) #Bellani
        ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=8) #Températures extrêmes
        ws.merge_cells(start_row=1, start_column=9, end_row=1, end_column=10) #Evaportation
        ws.merge_cells(start_row=1, start_column=11, end_row=3, end_column=11) #Pluies
        ws.merge_cells(start_row=1, start_column=12, end_row=1, end_column=16) #Température et Humidité de l'air à 6 heures
        ws.merge_cells(start_row=1, start_column=17, end_row=1, end_column=21) #Température et Humidité de l'air à 15 heures
        ws.merge_cells(start_row=1, start_column=22, end_row=1, end_column=26) #Température et Humidité de l'air à 18 heures
        ws.merge_cells(start_row=1, start_column=27, end_row=3, end_column=27) #Date
        # subheaders
        ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=7) #Abri
        ws.merge_cells(start_row=2, start_column=9, end_row=2, end_column=10) #Piche
        ws.merge_cells(start_row=2, start_column=12, end_row=2, end_column=16) #(Psychromètre a aspiration)
        ws.merge_cells(start_row=2, start_column=17, end_row=2, end_column=21) #(Psychromètre a aspiration)
        ws.merge_cells(start_row=2, start_column=22, end_row=2, end_column=26) #(Psychromètre a aspiration)


        # Label Date, Total and Average rows
        row_labels = ["1","2", "3", "4", "5", "Tot.", "Moy.", "6", "7", "8", "9", "10", "Tot.", "Moy.", "11", "12", "13", "14", "15", "Tot.", "Moy.", "16", "17", "18", "19", "20", "Tot.", "Moy.", "21", "22", "23", "24", "25", "Tot.", "Moy.", "26", "27", "28", "29", "30", "31", "Tot.", "Moy.", "Tot.", "Moy."]
        # Update the cells in the second and last column with the date values
        columns = [2, 27]
        for col in columns:
            for i, value in enumerate(row_labels, start=4):
                cell = ws.cell(row=i, column=col)
                cell.value = value
                # wb.save(new_version_of_file) # Save the modified workbook
        
        # Save Excel file
        file_path = os.path.join(save_dir, f'{method_name}_Excel_with_OCR_Results.xlsx')
        wb.save(file_path)
        results.append([file_path])

        
        ### FOR ILLUSTRATION PURPOSES: Uncomment the following lines if you want to see the ROIs per table
        # plt.imshow(ROIs_image)
        # plt.show()
    
    # Update paths to include the new Excel files (from the methods above) for merging. This is just a double check for correct cell placement
    top_excel_path = os.path.join(save_dir, 'Top_Excel_with_OCR_Results.xlsx')
    midpoint_excel_path = os.path.join(save_dir, 'Midpoint_Excel_with_OCR_Results.xlsx')
    path_to_save_merged_excel_file = os.path.join(pre_QA_QC_transcribed_hydroclimate_data_dir_station, f'{month_filename}_pre_QA_QC.xlsx')
    merge_excel_files(top_excel_path, midpoint_excel_path, path_to_save_merged_excel_file, 1, no_of_rows_including_headers) # this prioritizes the top coordinates of the bounding box to the mid point coordinates when placing the transcribed data into an excel sheet. But considers the best placement for both as double check. Here the 1 represent the first row and the 46 represents the last possible row to perform the merging of the excel files. In our case we have 43 rows with data + 3 header rows = 46.
    # merge_excel_files(midpoint_excel_path, top_excel_path, path_to_save_merged_excel_file, 1, no_of_rows_including_headers) # this prioritizes the midpoint coordinates of the bounding box to the top coordinates when placing the transcribed data into an excel sheet. But considers the best placement for both as double check. Here the 1 represent the first row and the 46 represents the last possible row to perform the merging of the excel files. In our case we have 43 rows with data + 3 header rows = 46.


    return path_to_save_merged_excel_file