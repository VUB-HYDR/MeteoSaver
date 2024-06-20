#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
#from paddleocr import PaddleOCR,draw_ocr
from PIL import Image, ImageDraw, ImageFont
import pytesseract
import easyocr
import tensorflow as tf
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
import math


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


def get_excel_cell_coordinates(file_path, sheet_name):
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]
    
    cell_coordinates = {}
    for row in ws.iter_rows(min_row=1, min_col=1, max_row=43, max_col=24):
        for cell in row:
            if cell.value is None:
                cell_coordinates[cell.coordinate] = (cell.column, cell.row)
                
    return cell_coordinates

def normalize_coordinates(coord, dimension):
    return coord / dimension

def normalize_cell_coordinates(cell_coordinates, sheet_dimensions):
    normalized_cell_coords = {}
    for cell_ref, (col, row) in cell_coordinates.items():
        normalized_col = normalize_coordinates(col, sheet_dimensions['width'])
        normalized_row = normalize_coordinates(row, sheet_dimensions['height'])
        normalized_cell_coords[cell_ref] = (normalized_col, normalized_row)
    return normalized_cell_coords

def find_closest_cell(normalized_center_x, normalized_center_y, normalized_cell_coordinates):
    closest_cell = None
    min_distance = float('inf')
    
    for cell_ref, (normalized_col, normalized_row) in normalized_cell_coordinates.items():
        distance = np.sqrt((normalized_center_x - normalized_col) ** 2 + (normalized_center_y - normalized_row) ** 2)
        
        if distance < min_distance:
            min_distance = distance
            closest_cell = cell_ref
    
    return closest_cell

# def find_next_empty_cell(ws, start_cell_ref):
#     column = start_cell_ref[0]
#     start_row = int(start_cell_ref[1:])
    
#     current_row = start_row
#     while ws[f"{column}{current_row}"].value is not None:
#         current_row += 1
    
#     return f"{column}{current_row}"

# def find_previous_cell(ws, start_cell_ref):
#     column = start_cell_ref[0]
#     start_row = int(start_cell_ref[1:])

#     # Check the very previous cell only
#     if start_row > 1:
#         next_row = start_row - 1

#         return f"{column}{next_row}"
#     else:
#         return f"{column}{start_row}"

def find_next_cell(ws, start_cell_ref):
    column = start_cell_ref[0]
    start_row = int(start_cell_ref[1:])

    # Check the very next cell only
    next_row = start_row + 1

    return f"{column}{next_row}"

# def find_adjacent_empty_cell(ws, start_cell_ref):
#     column = start_cell_ref[0]
#     start_row = int(start_cell_ref[1:])

#     # Check the cell directly below
#     current_row = start_row
#     if ws[f"{column}{current_row}"].value is not None:
#         current_row += 1

#     # Check the cell directly above
#     if current_row > 0 and ws[f"{column}{current_row}"].value is not None:
#         current_row -= 1
    
#     return f"{column}{current_row}"

    # # If both adjacent cells are occupied or out of bounds, return None
    # elseif: 
    #     return f"{column}{start_row}
def get_table_boundaries(contours):
    x_coords = []
    y_coords = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        x_coords.extend([x, x + w])
        y_coords.extend([y, y + h])
    return min(x_coords), min(y_coords), max(x_coords), max(y_coords)

# def calculate_cell_reference(center_x, center_y, max_rows, max_columns, table_width, table_height):
#     row = math.floor(center_y / table_height * max_rows) + 1
#     column = math.floor(center_x / table_width * max_columns) + 1

#     # Ensure the column and row indices are within valid ranges
#     if column < 1:
#         column = 1
#     elif column > max_columns:
#         column = max_columns

#     if row < 1:
#         row = 1
#     elif row > max_rows:
#         row = max_rows

#     return f'{openpyxl.utils.get_column_letter(column)}{row}'

def filter_contours(contours, min_width_threshold, min_height_threshold, max_width_threshold, max_height_threshold):
    filtered_contours = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if min_width_threshold <= w <= max_width_threshold and min_height_threshold <= h <= max_height_threshold:
            filtered_contours.append(contour)
    return filtered_contours

def calculate_cell_reference(center_x, row_index, max_columns, table_width):
    column = math.floor(center_x / table_width * max_columns) + 1

    # Ensure the column index is within valid ranges
    if column < 1:
        column = 1
    elif column > max_columns:
        column = max_columns

    return f'{openpyxl.utils.get_column_letter(column)}{row_index}'


def find_missing_y(y_positions, avg_height, image_height):
    y_positions = sorted(y_positions)
    missing_y = None
    for i in range(1, len(y_positions)):
        if y_positions[i] - y_positions[i - 1] > avg_height:
            missing_y = y_positions[i - 1] + avg_height
            break
    if not missing_y:
        missing_y = max(y_positions) + avg_height
        if missing_y + avg_height > image_height:
            missing_y = min(y_positions) - avg_height
    return missing_y

def group_contours_into_columns(contours, num_columns, image_width):
    column_width = image_width // num_columns
    columns = {i: [] for i in range(num_columns)}
    
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        column_index = min(x // column_width, num_columns - 1)
        columns[column_index].append((x, y, w, h))
    
    return columns

# def add_missing_rois(sorted_contours, space_threshold, width_threshold, max_height_est_per_box, num_columns, image_width):
#     # Group contours into columns
#     columns = group_contours_into_columns(sorted_contours, num_columns, image_width)

#     new_boxes = []
#     for i in sorted(columns.keys()):  # Ensure columns are processed in order
#         column_boxes = sorted(columns[i], key=lambda b: b[1])  # Sort by y-coordinate
#         column_count = len(column_boxes)
#         print(f'Number of current rows in the current column: {column_count}')  # Debug statement
#         for j in range(1, len(column_boxes)):
#             prev_box = column_boxes[j - 1]
#             curr_box = column_boxes[j]
#             space_between = curr_box[1] - (prev_box[1] + prev_box[3])
#             if space_between > space_threshold:
#                 # Clauculate the number of new boxes
#                 num_new_boxes = space_between // max_height_est_per_box
#                 if 0< num_new_boxes <= 1.6:
#                     new_y = prev_box[1] + prev_box[3] + (space_between - max_height_est_per_box) // 2
#                     new_height = max_height_est_per_box
#                     new_box = (prev_box[0]-10, new_y, width_threshold+10, new_height)
#                     print(f'Added new box at: {new_box}')  # Debug statement
#                     column_boxes.append(new_box)
#                 if num_new_boxes > 1.6:
#                     box_height = space_between // (num_new_boxes + 1)
#                     for k in range(1, num_new_boxes + 1):
#                         new_y = prev_box[1] + prev_box[3] + k * box_height - box_height // 2
#                         new_box = (prev_box[0]-10, new_y, width_threshold+10, box_height)
#                         print(f'Added new box at: {new_box}')  # Debug statement
#                         column_boxes.append(new_box)
#             # columns[x] = column_boxes
#         new_boxes.extend(column_boxes)
    
#     # new_boxes = [box for column in columns.values() for box in column]

#     new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32) for box in new_boxes]
#     # new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32).reshape((-1, 1, 2)) for box in new_boxes]
    

#     return new_contours

# def add_missing_rois(sorted_contours, space_threshold, width_threshold, max_height_est_per_box, num_columns, image_width):
#     # Group contours into columns
#     columns = group_contours_into_columns(sorted_contours, num_columns, image_width)

#     new_boxes = []
#     for i in sorted(columns.keys()):  # Ensure columns are processed in order
#         column_boxes = sorted(columns[i], key=lambda b: b[1])  # Sort by y-coordinate
#         column_count = len(column_boxes)
#         print(f'Number of current rows in the current column: {column_count}')  # Debug statement
#         for i in range(1, len(column_boxes)):
#             prev_box = column_boxes[i - 1]
#             curr_box = column_boxes[i]
#             space_between = curr_box[1] - (prev_box[1] + prev_box[3])
#             if space_between > space_threshold:
#                 # Calculate the y position for the new contour
#                 # new_y = prev_box[1] + prev_box[3]
#                 new_y = prev_box[1] + prev_box[3] + (space_between - max_height_est_per_box) // 2
#                 new_height = max_height_est_per_box
#                 new_box = (prev_box[0]-10, new_y, width_threshold+10, new_height)
#                 print(f'Added new box at: {new_box}')  # Debug statement
#                 column_boxes.append(new_box)
#             # columns[x] = column_boxes
#         new_boxes.extend(column_boxes)
    
#     # new_boxes = [box for column in columns.values() for box in column]

#     new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32) for box in new_boxes]
#     # new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32).reshape((-1, 1, 2)) for box in new_boxes]
    

#     return new_contours


def add_missing_rois(sorted_contours, space_threshold, width_threshold, max_height_est_per_box, max_rows, num_columns, image_width):
    # Group contours into columns
    columns = group_contours_into_columns(sorted_contours, num_columns, image_width)

    new_boxes = []
    for i in sorted(columns.keys()):  # Ensure columns are processed in order
        column_boxes = sorted(columns[i], key=lambda b: b[1])  # Sort by y-coordinate
        column_count = len(column_boxes)
        print(f'Number of current rows in the current column: {column_count}')  # Debug statement
        # Calculate gaps and sort them by size (largest first)
        gaps = []
        for j in range(1, len(column_boxes)):
            prev_box = column_boxes[j - 1]
            curr_box = column_boxes[j]
            space_between = curr_box[1] - (prev_box[1] + prev_box[3])
            if space_between > space_threshold:
                gaps.append((space_between, prev_box, curr_box))
        
        gaps.sort(key=lambda x: x[0], reverse=True)  # Sort gaps in between the cells by size (largest first)

        # Add new boxes for the gaps in priority order
        for gap in gaps:
            if column_count >= max_rows:
                break
            space_between, prev_box, curr_box = gap
            # Calculate the y position for the new contour
            new_y = prev_box[1] + prev_box[3] + (space_between - max_height_est_per_box) // 2
            new_height = max_height_est_per_box
            new_box = (prev_box[0]-10, new_y, width_threshold+10, new_height)
            print(f'Added new box at: {new_box}')  # Debug statement
            column_boxes.append(new_box)
            column_count += 1
            # columns[x] = column_boxes
        column_boxes = sorted(column_boxes, key=lambda b: b[1])  # Sort again after adding new boxes
        new_boxes.extend(column_boxes)
    
    # new_boxes = [box for column in columns.values() for box in column]

    new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32) for box in new_boxes]
    # new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32).reshape((-1, 1, 2)) for box in new_boxes]
    

    return new_contours


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
    # if ocr_model == 'PaddleOCR':
    #     ## Lauching PaddleOCR, which would be used by downloading necessary files as shown below
    #     paddle_ocr = PaddleOCR(use_angle_cls=True, lang = 'en', use_gpu=False) ## Run only once to download all required files
    if ocr_model == 'EasyOCR':
        ## Lauching EasyOCR
        easyocr_reader = easyocr.Reader(['en']) # this needs to run only once to load the model into memory

    contours = detected_table_cells[0]

    # Filter out smaller or larger bounding boxes from all the detected text contours. this is helpful to avoid overly large cells or small cells with no text
    min_width_threshold = 50
    min_height_threshold = 28
    max_width_threshold = 200
    max_height_threshold = 90
    
    filtered_contours = filter_contours(contours, min_width_threshold, min_height_threshold, max_width_threshold, max_height_threshold)

    image_with_all_bounding_boxes = detected_table_cells[1]
    table_copy = detected_table_cells[2]
    table_original_image = detected_table_cells[3]

    # Get the dimensions of the loaded image
    image_height, image_width, image_channels = image_with_all_bounding_boxes.shape
    
    min_x, min_y, max_x, max_y = get_table_boundaries(contours)
    table_width = max_x - min_x
    table_height = max_y - min_y

    ROIs_image = table_copy.copy()
    
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
        contours_sorted = sorted(filtered_contours, key=lambda c: cv2.boundingRect(c)[1])

         # Define the maximum number of rows, space threshold, and width for new ROIs
        max_rows = 43
        space_threshold = 50
        width_threshold = 120
        max_height_est_per_box = 50
        num_columns = 24  # Adjust this value based on your specific use case

        # Add missing ROIs to the contours
        new_contours = add_missing_rois(contours_sorted, space_threshold, width_threshold, max_height_est_per_box, max_rows, num_columns, image_width)

        # Organize the contours
        organized_rows = organize_method(new_contours, (image_height, image_width, image_channels), max_rows)
        
        sorted_rows = sorted(organized_rows, key=lambda row: np.median([cv2.boundingRect(c)[1] + cv2.boundingRect(c)[3] // 2 for c in row]))

        
        # # Use new_contours instead of original contours
        # all_contours = new_contours        # organized_rows = organize_and_merge_contours(contours_sorted, (image_height, image_width, image_channels), max_rows)
        
        
        # # # Define the row index to start filling from
        # # start_row_index = 1  # Change this to your desired starting row index

        # # Sort contours by y-coordinate
        # # contours_sorted = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

        # # Calculate the median y-coordinate for each row and sort rows by this median
        # sorted_rows = sorted(new_contours, key=lambda row: np.median([cv2.boundingRect(c)[1] + cv2.boundingRect(c)[3] // 2 for c in row]))

        for row_index, row in enumerate(sorted_rows, start=1):
            for contour in row:
        ## Text detection using an OCR model; Here using TesseractOCR
        # for contour in contours_sorted:

        #for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                

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

                # Draw the adjusted ROI on the output image
                cv2.rectangle(ROIs_image, (x, y), (x + w, y + h), (255, 0, 0), 2)  # Red color for ROI
                
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


                    # if ocr_model == 'PaddleOCR':
                    #     ## Using PaddleOCR
                    #     ocr_result = paddle_ocr.ocr('detected.png', cls = True)

                    if ocr_model == 'EasyOCR':
                    # Using EasyOCR
                        ocr_result = easyocr_reader.readtext('detected.png', detail = 0, allowlist='0123456789')
                    # Using OCR to recognize text/transcription
                    if ocr_result is not None:
                        
                        # Maximum number of columns and rows. These can be changed depending on the tables in the images
                        max_column_index = 24  # Number of columns in the table. Total number is original unclipped image are 27
                        max_row_index = 43  # Estimated number of rows in the table  .Previosly had it at 57 and results were good.       even 56 was good.     had this previously at 43
                        
                        # cell_width = max(image_width // max_column_index, min_width_threshold)
                        # cell_height = max(image_height//max_row_index, min_height_threshold)
                        
                        # # Track filled cells using a set
                        # filled_cells = []
                        
                        
                        # # Convert the x-coordinate to a column letter # Assuming x-coordinate translates to columns cell_column
                        # # Ensure x is within the valid range for Excel column indices
                        
                        # if 1 <= center_x <= image_width:  # Excel's maximum column index
                        #     column_index = int(max((center_x) // cell_width, 0)) + 1 # Ensure column index is at least 1
                        #     #cell_column = openpyxl.utils.get_column_letter(min(column_index, max_column_index))
                        # else:
                        #     column_index = 1
                        #     #cell_column = 'A'  # Set a default column if x is out of range
                        
                        # # if 1 <= center_y <= image_height:
                        # #     # Calculate the row index based on the row ratio
                        # #     row_index = (center_y / cell_height) + 1  # Calculate row index as a floating-point number
                            
                        # #     # Round the row index to the nearest integer
                        # #     cell_row = int(round(row_index))
                            
                        # #     # Ensure the row index is within the valid range
                        # #     cell_row = min(max(cell_row, 1), max_row_index)
                        # #     # row_ratio = (center_y) // cell_height
                        # #     # cell_row = min(int(row_ratio) + 1, max_row_index)  # Convert row ratio to integer and ensure it's within valid range
                        # #     # #cell_row = min(center_y // cell_height + 0.5, max_row_index)
                        # # else:
                        # #     cell_row = 1  # Set a default row if y is out of range  # Assuming y-coordinate translates to rows



                        # # Ms Excel Template cell coordinates
                        # file_path = f'docs\Table_structure.xlsx'
                        # sheet_name = 'clipped_tilted'
                        # cell_coordinates = get_excel_cell_coordinates(file_path, sheet_name)
                        # sheet_dimensions = {'width': max_column_index, 'height': max_row_index}  # Assuming max column index and max row index

                        # normalized_cell_coords = normalize_cell_coordinates(cell_coordinates, sheet_dimensions)


                        # # Find closest cell in template
                        # # Normalize bounding box center coordinates
                        # normalized_center_x = normalize_coordinates(center_x, image_width)
                        # normalized_center_y = normalize_coordinates(center_y, image_height)
                        # closest_cell = find_closest_cell(normalized_center_x, normalized_center_y, normalized_cell_coords)
        

                        # # if closest_cell:
                        # #     cell = ws[closest_cell]
                        # #     cell.value = ocr_result

                        # if closest_cell:
                        # # Check if the cell is occupied and find the next empty cell in the same column
                        #     cell_ref = closest_cell
                        #     while ws[cell_ref].value is not None:
                        #         cell_ref = find_next_cell(ws, cell_ref)
                            
                        #     cell = ws[cell_ref]
                        #     cell.value = ocr_result


                        # Calculate the cell reference
                        cell_ref = calculate_cell_reference(center_x, row_index, max_columns=24, table_width=table_width)
                        # Write the OCR result to the Excel cell
                        ws[cell_ref].value = ocr_result





                        # # Write the OCR value to the cell in the Excel sheet
                        # cell = ws.cell(row=row_index, column=column_index)


                        # # # Set the cell value to the OCR result
                        # # cell.value = ocr_result

                        # current_cell_value = cell.value

                        # # If the cell is empty, add the OCR result directly
                        # if current_cell_value is None:
                        #     cell.value = ocr_result
                        # else:
                        #     # If the cell already has content, append the new result
                        #     cell.value = f"{current_cell_value}{ocr_result}"

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


        # Insert two columns on the left side of the excel sheet
        ws.insert_cols(1, 2)
        
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

        plt.imshow(ROIs_image)
        plt.show()

    return results