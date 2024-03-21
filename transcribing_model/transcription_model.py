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




def transcription(detected_table_cells, ocr_model):
    '''
    # Makes use of a pre-processed image (in grayscale) to detect the table from the record sheets

    Parameters
    --------------
    detected_table_cells where: 
        detected_table_cells[0]: contours. Contours for the detected text in the table cells
        detected_table_cells[1]: image_with_all_bounding_boxes. Bounding boxex for each cell for which clips will be made later before optical character recognition 
        detected_table_cells[2]: table_copy

    ocr_model: Optical Character Recognition/Handwritten Text Recognition of choice
    
    Returns
    -------------- 
    wb : MS Excel workbook with OCR Results
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

    # Get the dimensions of the loaded image
    image_height, image_width, image_channels = image_with_all_bounding_boxes.shape

    ## Create an Excel workbook and add a worksheet where the transcribed text will be saved
    wb = Workbook()
    ws = wb.active
    ws.title = 'OCR_Results'

    # Define the row index to start filling from
    start_row_index = 0  # Change this to your desired starting row index

    # Sort contours by y-coordinate
    contours_sorted = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

    # Initialize current_row_index to start from the specified index
    current_row_index = start_row_index

    ## Text detection using an OCR model; Here using TesseractOCR
    for contour in contours_sorted:
    #for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        
        # Adjust these threshold values according to your requirements
        min_width_threshold = 20
        min_height_threshold = 10  # Had this at 10 previously

        # Filter out smaller bounding boxes
        if w >= min_width_threshold and h >= min_height_threshold:
            
            # # Calculate a factor to increase the bounding box area (e.g., 80% larger)
            # increase_factor_width = 0.1  # Modify this factor as needed
            # increase_factor_height = 0.5 # Modify this factor as needed
        
            # x -= int(w * increase_factor_width)  # Increase width
            # y -= int(h * increase_factor_height)  # Increase height
            # w += int(2 * w * increase_factor_width)  # Increase width
            # h += int(2 * h * increase_factor_height)  # Increase height
            
            
            #bounding_boxes.append((x, y, w, h))
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
                    ocr_result = pytesseract.image_to_string('detected.png', lang='cc-base', config='--psm 7 -c tessedit_char_whitelist=0123456789.')
                    #ocr_result = pytesseract.image_to_string('detected.png', lang='cobedore-V6+engold', config='--psm 7 -c tessedit_char_whitelist=0123456789.')

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
                    max_row_index = 57  # Estimated number of rows in the table  . Initital runs had 56             
                    
                    cell_width = max(image_width // max_column_index, min_width_threshold)
                    cell_height = max(image_height//max_row_index, min_height_threshold)
                    
                    # Track filled cells using a set
                    filled_cells = []
                    
                    
                    # Convert the x-coordinate to a column letter # Assuming x-coordinate translates to columns cell_column
                    # Ensure x is within the valid range for Excel column indices
                    
                    if 1 <= center_x <= image_width:  # Excel's maximum column index
                        column_index = max((center_x) // cell_width, 0) + 1 # Ensure column index is at least 1
                        #cell_column = openpyxl.utils.get_column_letter(min(column_index, max_column_index))
                    else:
                        column_index = 0
                        #cell_column = 'A'  # Set a default column if x is out of range
                    

                    # # Increment current_row_index for each new row
                    # if center_y // cell_height + 1 > current_row_index:
                    #     current_row_index = center_y // cell_height + 1
                        
                    # Calculate row index based on y-coordinate
                    row_index = current_row_index + int((center_y - 0) / cell_height)


                    # if 1 <= center_y <= image_height:
                    #     # Calculate the row index based on the row ratio
                    #     row_index = (center_y / cell_height) + 1  # Calculate row index as a floating-point number
                        
                    #     # Round the row index to the nearest integer
                    #     cell_row = round(row_index)
                        
                    #     # Ensure the row index is within the valid range
                    #     cell_row = min(max(cell_row, 1), max_row_index)
                    #     # row_ratio = (center_y) // cell_height
                    #     # cell_row = min(int(row_ratio) + 1, max_row_index)  # Convert row ratio to integer and ensure it's within valid range
                    #     # #cell_row = min(center_y // cell_height + 0.5, max_row_index)
                    # else:
                    #     cell_row = 1  # Set a default row if y is out of range  # Assuming y-coordinate translates to rows
                
                    
                    # cell_key = (cell_row, column_index)      
                    # # Check if the cell is already filled
                    # if cell_key in filled_cells:
                    #     cell_row = cell_row + 1  # Move to the next row if the cell is filled
                    #     cell_key = (cell_row, column_index)  # Update the cell key
                    #     filled_cells.append(cell_key)  # Add the filled cell coordinates to the set
                    # else:
                    #     filled_cells.append(cell_key)  # Still add the filled cell coordinates to the set

                    # Write the OCR value to the cell in the Excel sheet
                    cell = ws.cell(row=row_index, column=column_index)
                    #cell = ws.cell(row=cell_row, column=column_index)

                    # Check if the cell is already populated
                    if cell.value:
                        # If the cell already contains text, append the new text with a newline character
                        cell.value += '\n' + ocr_result
                    else:
                        # If the cell is empty, set the cell value to the new text
                        cell.value = ocr_result

                    #cell.value = ocr_result

                    # Increment row index for the next iteration
                    current_row_index = max(current_row_index, row_index + 1)

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
            
    wb.save('Excel_with_OCR_Results.xlsx') 

    return wb

