#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
from paddleocr import PaddleOCR,draw_ocr
from PIL import Image, ImageDraw, ImageFont
import pytesseract
import easyocr
import tensorflow as tf
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans


def organize_contours(contours, image_shape, max_rows):
    '''
    # Organizes the bounding boxes, here termed as contours, in rows by their center co-ordinates using the KMeans clustering

    Parameters
    --------------
    contours: List of contours for the detected text in the table cells with coordinates
    image_shape: Image of the table
    max_rows: Maximum rows, adjust based on your table's expected structure

    Returns
    -------------- 
    rows: Bounding boxes organised in rows using Kmeans clustering

    '''

    image_height, image_width, _ = image_shape
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

# def organize_contours_bottom(contours, image_shape, max_rows):
    # '''
    # # Organizes the bounding boxes, here termed as contours, in rows by their bottom co-ordinates using the KMeans clustering

    # Parameters
    # --------------
    # contours: List of contours for the detected text in the table cells with coordinates
    # image_shape: Image of the table
    # max_rows: Maximum rows, adjust based on your table's expected structure

    # Returns
    # -------------- 
    # rows: Bounding boxes organised in rows using Kmeans clustering

    # '''

#     image_height, image_width, _ = image_shape
#     bottom = [(cv2.boundingRect(contour)[1] + cv2.boundingRect(contour)[3]) for contour in contours]
#     if len(bottom) == 0:
#         return []
#     kmeans = KMeans(n_clusters=min(max_rows, len(bottom)), random_state=0)
#     kmeans.fit(np.array(bottom).reshape(-1, 1))
#     labels = kmeans.labels_

#     rows = [[] for _ in range(max_rows)]
#     for label, contour in zip(labels, contours):
#         rows[label].append(contour)
    
#     for i in range(len(rows)):
#         rows[i] = sorted(rows[i], key=lambda c: cv2.boundingRect(c)[0])
    
#     return rows


def organize_contours_top(contours, image_shape, max_rows):
    '''
    # Organizes the bounding boxes, here termed as contours, in rows by their top co-ordinates using the KMeans clustering

    Parameters
    --------------
    contours: List of contours for the detected text in the table cells with coordinates
    image_shape: Image of the table
    max_rows: Maximum rows, adjust based on your table's expected structure

    Returns
    -------------- 
    rows: Bounding boxes organised in rows using Kmeans clustering

    '''

    image_height, image_width, _ = image_shape
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




def transcription(detected_table_cells, ocr_model):
    '''
    # Makes use of a pre-processed image (in grayscale) to detect the table from the record sheets

    Parameters
    --------------
    detected_table_cells where: 
        detected_table_cells[0]: contours. Contours for the detected text in the table cells
        detected_table_cells[1]: image_with_all_bounding_boxes. Bounding boxex for each cell for which clips will be made later before optical character recognition 
        detected_table_cells[2]: table_copy
        detected_table_cells[3]: table_original_image
    

    ocr_model: Optical Character Recognition/Handwritten Text Recognition of choice
    
    Returns
    -------------- 
    wb : Ms Excel workbook with OCR Results
    ''' 
    if ocr_model == 'Tesseract-OCR':
        ## Lauching Tesseract-OCR
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' ## Here input the PATH to the Tesseract executable on your computer. See more information here: https://pypi.org/project/pytesseract/
    if ocr_model == 'PaddleOCR':
        ## Lauching PaddleOCR, which would be used by downloading necessary files as shown below
        paddle_ocr = PaddleOCR(use_angle_cls=True, lang = 'en', use_gpu=False) ## Run only once to download all required files
    if ocr_model == 'EasyOCR':
        ## Lauching EasyOCR
        easyocr_reader = easyocr.Reader(['en']) # this needs to run only once to load the model into memory

    contours = detected_table_cells[0]
    image_with_all_bounding_boxes = detected_table_cells[1]
    table_copy = detected_table_cells[2]
    table_original_image = detected_table_cells[3]

    # Get the dimensions of the loaded image
    image_height, image_width, image_channels = image_with_all_bounding_boxes.shape

    results = []

    organize_methods = {
        #'Bottom': organize_contours_bottom,
        'Midpoint': organize_contours,
        'Top': organize_contours_top
    }

    for method_name, organize_method in organize_methods.items():

        ## Create an Excel workbook and add a worksheet where the transcribed text will be saved
        wb = Workbook()
        ws = wb.active
        ws.title = 'OCR_Results'

        # Sort contours by y-coordinate
        contours_sorted = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

        max_rows = 45  # maximum rows, adjust based on your table's expected structure. Here, I had started with 45 rows and it was already giving good results 
        organized_rows = organize_method(contours_sorted, (image_height, image_width, image_channels), max_rows)
        # organized_rows = organize_and_merge_contours(contours_sorted, (image_height, image_width, image_channels), max_rows)


        # # Define the row index to start filling from
        # start_row_index = 1  # Change this to your desired starting row index

        # Sort contours by y-coordinate
        # contours_sorted = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

        # Calculate the median y-coordinate for each row and sort rows by this median
        sorted_rows = sorted(organized_rows, key=lambda row: np.median([cv2.boundingRect(c)[1] + cv2.boundingRect(c)[3] // 2 for c in row]))

        for row_index, row in enumerate(sorted_rows, start=1):
            for contour in row:
        ## Text detection using an OCR model; Here using TesseractOCR
        # for contour in contours_sorted:

        #for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                
                # Adjust these threshold values according to your requirements
                min_width_threshold = 20 # 50 previously
                min_height_threshold = 28  # Had this at 13 previously. 20, 25, 30. trying 28, but it might be over fitting 

                max_width_threshold = 200 # 120 previously
                max_height_threshold = 90 # 60 previously

                # Filter out smaller bounding boxes
                if (min_width_threshold <= w <= max_width_threshold and min_height_threshold <= h <= max_height_threshold):

                    # Calculate a factor to increase the bounding box area (e.g., 80% larger)
                    factor_width = 0.05  # Modify this factor as needed
                    increase_factor_height = 0.25 # Modify this factor as needed
                
                    x += int(w * factor_width)  # Increase width
                    y -= int(h * increase_factor_height)  # Increase height
                    w -= int(1 * w * factor_width)  # Decrease width a little to avoid vertical lines that may be transcribed as the number 1 yet they aren't a number
                    h += int(2 * h * increase_factor_height)  # Increase height
                    
                    # This line below is about
                    # drawing a rectangle on the image with the shape of
                    # the bounding box. Its not needed for the OCR.
                    # Its just added for debugging purposes.
                    image_with_all_bounding_boxes = cv2.rectangle(image_with_all_bounding_boxes, (x, y), (x + w, y + h), (0, 255, 0), 5)
                    
                    # Calculate center coordinates of the bounding box
                    center_x = x + w // 2
                    center_y = y + h // 2
                    
                    # OCR
                    # Crop each cell using the bounding rectangle coordinates
                    ROI = table_copy[y:y+h, x:x+w]
                    
                    if ROI.size != 0:  # Check if the height and width are greater than zero
                        
                        cv2.imwrite('detected.png', ROI)
                        if ocr_model == 'Tesseract-OCR':
                        # Using Tesseract-OCR
                            ocr_result = pytesseract.image_to_string('detected.png', lang='cobedore-V6', config='--psm 7 -c tessedit_char_whitelist=0123456789.') # Just added -c tessedit_char_whitelist=0123456789. to really limit the text detected
                            #ocr_result = pytesseract.image_to_string('detected.png', lang='ccc-base+eng+cobedore-V6', config='--psm 7 -c tessedit_char_whitelist=0123456789.')

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


                        if ocr_model == 'PaddleOCR':
                            ## Using PaddleOCR
                            ocr_result = paddle_ocr.ocr('detected.png', cls = True)

                        if ocr_model == 'EasyOCR':
                        # Using EasyOCR
                            ocr_result = easyocr_reader.readtext('detected.png', detail = 0, allowlist='0123456789')
                        # Using OCR to recognize text/transcription
                        if ocr_result is not None:
                            
                            # Maximum number of columns and rows. These can be changed depending on the tables in the images
                            max_column_index = 27  # Number of columns in the table. 
                            max_row_index = 45  # Estimated number of rows in the table  .Previosly had it at 57 and results were good.       even 56 was good
                            
                            cell_width = max(image_width // max_column_index, min_width_threshold)
                            cell_height = max(image_height//max_row_index, min_height_threshold)
                            
                            # Track filled cells using a set
                            filled_cells = []
                            
                            
                            # Convert the x-coordinate to a column letter # Assuming x-coordinate translates to columns cell_column
                            # Ensure x is within the valid range for Excel column indices
                            
                            if 1 <= center_x <= image_width:  # Excel's maximum column index
                                column_index = int(max((center_x) // cell_width, 0)) + 1 # Ensure column index is at least 1
                                #cell_column = openpyxl.utils.get_column_letter(min(column_index, max_column_index))
                            else:
                                column_index = 1
                                #cell_column = 'A'  # Set a default column if x is out of range
                            
                            # if 1 <= center_y <= image_height:
                            #     # Calculate the row index based on the row ratio
                            #     row_index = (center_y / cell_height) + 1  # Calculate row index as a floating-point number
                                
                            #     # Round the row index to the nearest integer
                            #     cell_row = int(round(row_index))
                                
                            #     # Ensure the row index is within the valid range
                            #     cell_row = min(max(cell_row, 1), max_row_index)
                            #     # row_ratio = (center_y) // cell_height
                            #     # cell_row = min(int(row_ratio) + 1, max_row_index)  # Convert row ratio to integer and ensure it's within valid range
                            #     # #cell_row = min(center_y // cell_height + 0.5, max_row_index)
                            # else:
                            #     cell_row = 1  # Set a default row if y is out of range  # Assuming y-coordinate translates to rows


                            # Write the OCR value to the cell in the Excel sheet
                            cell = ws.cell(row=row_index, column=column_index)

                            # # Check if the cell is already populated
                            # while cell.value is not None:
                            #     # Move to the next row if the current cell is filled
                            #     row_index += 1
                            #     cell = ws.cell(row=row_index, column=column_index)

                            # Set the cell value to the OCR result
                            cell.value = ocr_result

                            # Set up border styles for excel output
                            thin_border = Border(
                                left=Side(style='thin'),
                                right=Side(style='thin'),
                                top=Side(style='thin'),
                                bottom=Side(style='thin'))

                            # Loop through cells to apply borders
                            for row in ws.iter_rows(min_row=1, max_row=max_row_index, min_col=1, max_col=max_column_index):
                                for cell in row:
                                    cell.border = thin_border
                            
                        else:
                            print('No values detected in clip')
                    else:
                        print('ROI is empty or invalid')


        # Insert a new row at the top for headers
        ws.insert_rows(1, amount = 3)

        # # Merging cells for headers
        # # Merge cells for headers that span multiple columns
        # ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=8)  # For "Températures extrêmes"
        # #*** Do this for the other cells that need to be merged


        # Define your headers (adjust as needed)
        headers = ["No de la pentade", "Date", "Bellani (gr. Cal/cm2) 6-6h", "Températures extrêmes", "", "", "", "", "Evaportation en cm3 6 - 6h", "", "Pluies en mm. 6-6h", "Température et Humidité de l'air à 6 heures", "", "", "", "", "Température et Humidité de l'air à 15 heures",  "", "", "", "", "Température et Humidité de l'air à 18 heures",  "", "", "", "", "Date"]
        sub_headers_1 = ["", "", "", "Abri", "", "", "", "", "Piche", "", "", "(Psychromètre a aspiration)", "", "", "", "", "(Psychromètre a aspiration)", "", "", "", "", "(Psychromètre a aspiration)", "", "", "", ""]
        sub_headers_2 =["", "", "", "Max.", "Min.", "(M+m)/2", "Ampl.", "Min. gazon", "Abri.", "Ext.", "", "T", "T'a", "e.", "U", "∆e", "T", "T'a", "e.", "U", "∆e","T", "T'a", "e.", "U", "∆e", ""]

        # Add the headers to the first row
        for col_num, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_num, value=header)
        # Add the first row of sub-headers to the second row
        for col_num, sub_header in enumerate(sub_headers_1, start=1):
            ws.cell(row=2, column=col_num, value=sub_header)

        # Add the second row of sub-headers to the third row
        for col_num, sub_header in enumerate(sub_headers_2, start=1):
            ws.cell(row=3, column=col_num, value=sub_header)

        file_path = f'src\output\{method_name}_Excel_with_OCR_Results.xlsx'
        wb.save(file_path)
        results.append([file_path])

        #wb.save(f'{method_name}_Excel_with_OCR_Results.xlsx') 

        # plt.imshow(image_with_all_bounding_boxes)
        # plt.show()

    return results


