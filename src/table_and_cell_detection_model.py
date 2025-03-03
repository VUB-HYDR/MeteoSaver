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


def filter_contours(contours, min_width_threshold, min_height_threshold, max_width_threshold, max_height_threshold):
    '''
    Filters a list of contours based on specified width and height thresholds.

    This function takes a list of contours and filters them by their bounding rectangle dimensions. Contours whose bounding rectangles fall within the specified width and height thresholds are retained, while others are discarded. This is useful for removing noise or irrelevant contours based on size constraints e.g. small contours that don't contain text or very large contours that cover more than one cell.

    Parameters
    --------------
    contours : list
        A list of contours, where each contour is an array of points defining the contour's shape.
    
    min_width_threshold : int
        The minimum width (in pixels) that a contour's bounding rectangle must have to be included in the filtered results.
    
    min_height_threshold : int
        The minimum height (in pixels) that a contour's bounding rectangle must have to be included in the filtered results.
    
    max_width_threshold : int
        The maximum width (in pixels) that a contour's bounding rectangle can have to be included in the filtered results.
    
    max_height_threshold : int
        The maximum height (in pixels) that a contour's bounding rectangle can have to be included in the filtered results.

    Returns
    --------------
    filtered_contours : list of bounding boxes (filtered)
        A list of contours that meet the specified width and height criteria. Contours whose bounding rectangles do not fall within the given thresholds are excluded.
    '''

    filtered_contours = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if min_width_threshold <= w <= max_width_threshold and min_height_threshold <= h <= max_height_threshold:
            filtered_contours.append(contour)
    return filtered_contours


def group_contours_into_columns(contours, num_columns, image_width):
    '''
    Groups contours into columns based on their horizontal position within an image.

    This function organizes a list of contours into a specified number of columns by calculating the column index for each contour based on its x-coordinate. The image is divided into equal-width columns, and each contour is assigned to a column based on where its bounding box falls horizontally. The resulting groups of contours are returned as a dictionary where the keys represent column indices.

    Parameters
    --------------
    contours :  list 
        A list of contours, where each contour is an array of points that define the contour's shape.

    num_columns : int
        The number of columns (from the expected table structure) into which the contours should be grouped. This value determines how the image width is divided.

    image_width : int
        The total width of the image/table (in pixels). This is used to calculate the width of each column and to determine in which column each contour belongs.

    Returns
    --------------
    columns : dict
        A dictionary where each key is a column index (ranging from 0 to `num_columns` - 1), and the value is a list of tuples. Each tuple represents a contour's bounding box in the format `(x, y, w, h)`, indicating its position and size within the image.
    '''

    columns = {i: [] for i in range(num_columns)}
    column_width = image_width // num_columns
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        column_index = min(x // column_width, num_columns - 1)
        columns[column_index].append((x, y, w, h))
    return columns


def add_missing_rois(sorted_contours, space_threshold, width_threshold, max_height_per_box, max_rows, num_columns, image_width):
    '''
    Identifies and adds missing regions of interest (ROIs) in a table by filling gaps between detected contours within each column.

    This function takes a set of sorted contours representing detected table cells and analyzes gaps between them within each column. If significant gaps are found, it adds new bounding boxes (ROIs) in those gaps to ensure that all expected rows are accounted for. This process helps to detect and fill in any missing cells that were not initially identified during contour detection. The newly generated contours are returned for further processing.

    Parameters
    --------------
    sorted_contours : list
        A list of contours sorted in the desired order (typically by their y-coordinate within the table). These contours represent the detected cells in the table.

    space_threshold : int
        The minimum vertical space (in pixels) between two consecutive contours within a column that should be considered a gap. If the space exceeds this threshold, a new ROI is added to fill the gap.

    width_threshold : int
        The width (in pixels) of the new ROI boxes that will be added to fill the gaps. This ensures that the new boxes have a consistent width relative to the other cells in the column.

    max_height_per_box : int
        The estimated maximum height (in pixels) for the new ROI boxes. This value is used to determine the size of the gaps and how to place the new boxes.

    max_rows : int
        The maximum number of rows expected in each column (from the expected table structure). The function will not add more boxes than this number, ensuring that the final column does not exceed the expected row count.

    num_columns : int
        The number of columns in the table. This determines how the contours are grouped and processed.

    image_width : int
        The total width of the image/table (in pixels). This is used to calculate the width of each column and group contours accordingly.

    Returns
    --------------
    new_contours : list
        A list of new contours representing the updated set of detected and added ROIs. These contours include both the original contours and any new ones created to fill gaps.
    '''


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
            new_y = prev_box[1] + prev_box[3] + (space_between - max_height_per_box) // 2
            new_height = max_height_per_box
            new_box = (prev_box[0], new_y, width_threshold+50, new_height) # extra 50 to esnure full capture of cell

            # Check for neighboring boxes within the same row
            has_left_neighbor = any(abs(prev_box[0] - b[0]) <= width_threshold for b in column_boxes)
            has_right_neighbor = any(abs(curr_box[0] - b[0]) <= width_threshold for b in column_boxes)

            if has_left_neighbor or has_right_neighbor:
                print(f'Added new box at: {new_box}')  # Debug statement
                column_boxes.append(new_box)
                column_count += 1

        column_boxes = sorted(column_boxes, key=lambda b: b[1])  # Sort again after adding new boxes
        new_boxes.extend(column_boxes)
    
    new_contours = [np.array([[box[0], box[1]], [box[0] + box[2], box[1]], [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]], dtype=np.int32) for box in new_boxes]
    
    return new_contours

# --- Noise Reduction ---
def reduce_noise(image):
    """Apply morphological operations to reduce small dots and noise."""
    kernel = np.ones((2,2), np.uint8)  # Adjust the kernel size as needed
    return cv2.morphologyEx(image, cv2.MORPH_OPEN, kernel, iterations=1)



# def remove_vertical_lines(image):
#     """Detect and remove only long vertical lines while keeping handwritten text intact."""
#     # Step 1: Convert to binary (adaptive thresholding)
#     binary = cv2.adaptiveThreshold(image, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 8)

#     # Step 2: Detect vertical lines using morphology
#     vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 50))  # Tall kernel to detect vertical lines
#     vertical_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)

#     # Step 3: Subtract detected vertical lines from the binary image
#     no_lines = cv2.bitwise_and(binary, cv2.bitwise_not(vertical_mask))

#     # Step 4: Restore broken text by slightly dilating it
#     text_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 2))
#     restored_text = cv2.dilate(no_lines, text_kernel, iterations=1)

#     # Step 5: Uninvert the image to maintain original black text on white background
#     final_output = cv2.bitwise_not(restored_text)

#     return final_output


def remove_vertical_lines(image):
    """Detect and remove only long vertical lines and small dots while keeping handwritten text intact."""
    # Step 1: Convert to binary (adaptive thresholding)
    binary = cv2.adaptiveThreshold(image, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 8)

    # Step 2: Detect vertical lines using morphology
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 50))  # Tall kernel to detect vertical lines
    vertical_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)

    # Step 3: Subtract detected vertical lines from the binary image
    no_lines = cv2.bitwise_and(binary, cv2.bitwise_not(vertical_mask))

    # Step 4: Find contours (text + unwanted dots)
    contours, _ = cv2.findContours(no_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Step 5: Remove small dots based on area threshold
    filtered_image = no_lines.copy()
    # for contour in contours:
    #     x, y, w, h = cv2.boundingRect(contour)
        
    #     # Define size threshold: Remove very small dots (too small to be text)
    #     if w * h < 70:  # Adjust based on dataset
    #         cv2.drawContours(filtered_image, [contour], -1, (0, 0, 0), thickness=cv2.FILLED)

    # Step 6: Apply **vertical-only dilation** to repair text (without making it bold)
    text_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 2))  # Expands only vertically
    restored_text = cv2.dilate(filtered_image, text_kernel, iterations=1)

    # Step 7: Uninvert the image to maintain black text on a white background
    final_output = cv2.bitwise_not(restored_text)

    return final_output

def table_and_cell_detection(image_in_grayscale, binarized_image, original_image, station, month_filename, transient_transcription_output_dir, clip_up, clip_down, clip_left, clip_right, max_table_width, max_table_height, min_cell_width_threshold, min_cell_height_threshold, max_cell_width_threshold, max_cell_height_threshold, space_height_threshold, space_width_threshold, max_cell_height_per_box, no_of_rows, no_of_columns):
    '''
    Detects and extracts tables and cells from a grayscale image using a combination of OpenCV-based image processing techniques.
    
    This function identifies the largest table-like region in the provided image, clips it using specified margins, and isolates table cells 
    by removing horizontal and vertical lines. It provides robust handling of cases where automatic table detection fails by allowing for 
    manual clipping. The processed table is then used for further text detection and analysis.

    Parameters
    --------------
    image_in_grayscale : numpy.ndarray
        The pre-processed grayscale version of the original image, used for table detection.
    binarized_image : numpy.ndarray
        The binarized version of the grayscale image where pixel intensities are reduced to binary values (0 or 255).
    original_image : numpy.ndarray
        The original colored input image for extracting the table in its original form.
    station : str
        Identifier of the station (station no.), used for saving outputs.
    month_filename : str
        Name of the file being processed, representing the specific month and year.
    transient_transcription_output_dir : str
        Directory to save intermediate results.
    clip_up, clip_down, clip_left, clip_right : int
        Number of pixels to clip from each side of the detected table: (i) from the top of the detected table for removing headers, (ii) from the bottom of the detected table for excluding unnecessary bottom parts of the table, (iii) from the left side of the detected table typically for removing row labels (pentad no and date since these are repetitive), and (iv) from the right side of the detected table, usually for excluding excess margins and the extra date column.
    max_table_width, max_table_height : int
        Maximum allowable dimensions for the detected table. If exceeded, manual clipping is applied.
    min_cell_width_threshold, min_cell_height_threshold : int
        Minimum width and height (in pixels) for valid table cells.
    max_cell_width_threshold, max_cell_height_threshold : int
        Maximum width and height (in pixels) for valid table cells.
    space_height_threshold, space_width_threshold : int
        Threshold for vertical and horizontal spacing between bounding boxes to identify missing cells.
    max_cell_height_per_box : int
        Maximum height allowed for any single cell.
    no_of_rows, no_of_columns : int
        Expected number of rows and columns in the table, used for detecting and filling missing cells.

    Returns
    --------------
    detected_table_cells : list
        A list containing:
        - detected_table_cells[0]: contours representing the detected text in table cells.
        - detected_table_cells[1]: image with bounding boxes drawn around detected table cells.
        - detected_table_cells[2]: binarized version of the detected table.
        - detected_table_cells[3]: clipped original table image.
        - detected_table_cells[4]: full detected table (unclipped), including headers and row labels.


    Notes
    --------------
    - If no table is detected via automatic contour detection, manual clipping is applied based on known table dimensions.
    - The function also removes horizontal and vertical lines, including dotted lines, to isolate text in the table cells.
    - The dimensions of the table are customizable based on the dataset used, and clipping values can be set to 0 to keep the full table.
    - Error handling is included to return None if table detection fails for a specific station and month.
    - The function supports visualizations for debugging by uncommenting the relevant sections of the code.

    '''


    # Here, we employ ML algorithms from Open Source Computer Vision (OpenCV) following methodologies similar to those described in https://livefiredev.com/how-to-extract-table-from-image-in-python-opencv-ocr/ [GitHub repository: \url{https://github.com/livefiredev/ocr-extract-table-from-image-python}, (last access: 19 July 2024)], but further customizing them for our case study.
    ## Using the pre-processing image to detect the table from the record sheets
    # Here, the threshold value for pixel intensities = 0, and the value 255 is assigned if the pixel value is above the threshold
    thresh = cv2.threshold(image_in_grayscale,0,255,cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1] 
    # Perform morphological operations (like dilation and erosion) for better segmentation
    kernel = np.ones((5,5),np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # Minimum dimensions of table
    threshold_area = 4  # Minimum contour area to consider as a cell
    threshold_height = 2   # Minimum height of the cell
    threshold_width = 2    # Minimum width of the cell
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
    ## Only for visualization purposes 
    # # Draw bounding box for the largest contour (if found), which here represents the table on the record sheets
    # if largest_contour is not None:
    #     x, y, w, h = cv2.boundingRect(largest_contour)
    #     cv2.rectangle(original_image, (x, y), (x + w, y + h), (0, 255, 0), 2)
    #     full_detected_table_with_labels = original_image[y:y + h, x:x + w] 
    #     table = image_in_grayscale[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
    #     #table = deskew(table) # Deskew the image, # Optional: Incase some of your images are skewed.
    #     table_original_image = image_in_grayscale[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
    #     # cv2.imwrite('table_original_image.jpg', table_original_image)
    # else:
    #     table = original_image # Incase the main table is not detected as the largest contour, we just use the original image/ whole record sheet as the image with the table
    #     table_original_image = original_image
    #     full_detected_table_with_labels = cv2.adaptiveThreshold(table, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 91,6) # Thresholding to reduce the image to black or white pixels
    #     # cv2.imwrite('table_original_image.jpg', table_original_image)

    # Draw bounding box for the largest contour (if found), which here represents the table on the record sheets
    if largest_contour is not None:
        x, y, w, h = cv2.boundingRect(largest_contour)
        cv2.rectangle(binarized_image, (x, y), (x + w, y + h), (0, 255, 0), 2)
        full_detected_table_with_labels = binarized_image[y:y + h, x:x + w]

        # Check if the table image dimensions exceed the thresholds to avoid sheets without proper table detection. This is customizable for different sheets. In our case, we had one type of sheets and an approximate uniform sheet dimensions
        height, width = full_detected_table_with_labels.shape[:2]
        if width <= max_table_width and height <= max_table_height: # These average table dimensions (in pixels; ~3900x3600) were determined from our sample sheets in the dataset given we followed similar protocol to digitize (image/scan) the data sheets.
            # These are therefore the AUTO-DETECTED TABLES using openCV 
            table = binarized_image[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
            
            # table = deskew(table) # Deskew the image, # Optional, uncomment if you'd like to use this: Incase some of your images are skewed.
            
            table_original_image = original_image[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
            # cv2.imwrite('table_original_image.jpg', table_original_image)

        else: # This indicates that the actual table was not detected from the image rather the whole sheet as a the table (for example, due to thick page boarders detected as a table)
            # We there make use of the knowledge of the average table dimensions (in pixels) in relation to the images that were determined from our sample sheets in the dataset to determine location of the table
            # This is thus a bug fix i.e. the MANUAL alternative to the table detection, where the AUTO-DETECTION does not detect the actual table.
            # Calculate the amount to clip from each side
            clip_x = (width - max_table_width) // 2  # Approximate table width in pixels = 3900. # Adjust these values according to your table 
            clip_y = (height - max_table_height) // 2 # Approximate table height in pixels = 3600. # Adjust these values according to your table 

            table = binarized_image[y + clip_y + 630:y + h - clip_y - 250, x + clip_x + 350:x + w - clip_x - 180]  # Here we manually clip the sheets to ensure clipping of the HEADERS and ROW LABELS (Date & Pentad no. in our case) from the table (table detected manually). Adjust this to your case study.
            table_original_image = original_image[y + clip_y + 630:y + h - clip_y - 250, x + clip_x + 350:x + w - clip_x - 180]
            
            
    else:
        # If no largest contour is detected. This indicates that the NO table was not detected from the image. Therefore we use the  entire image and make use of the knowledge of the average table dimensions (in pixels) in relation to the images that were determined from our sample sheets in the dataset to determine location of the table.
        # This is thus the MANUAL alternative to the table detection.
        height, width = image_in_grayscale.shape[:2]
        x, y, w, h = 0, 0, width, height  # Consider the entire image dimensions

        # Calculate the amount to clip from each side
        clip_x = (width - max_table_width) // 2
        clip_y = (height - max_table_height) // 2

        table = binarized_image[y + clip_y + 630:y + h - clip_y - 250, x + clip_x + 350:x + w - clip_x - 180]  # Here we manually clip the sheets to ensure clipping of the HEADERS and ROW LABELS (Date & Pentad no. in our case) from the table (table detected manually). Adjust this to your case study.
        table_original_image = original_image[y + clip_y + 630:y + h - clip_y - 250, x + clip_x + 350:x + w - clip_x - 180]
        # table = preprocessed_image[clip_y + clip_up:h - clip_y - clip_down, clip_x + clip_left:w - clip_x - clip_right] # Incase the main table is not detected as the largest contour, we just use the original image/ whole record sheet as the image with the table and clip it to manually set dimensions. These could have to be user input
        # table_original_image = original_image[clip_y + clip_up:h - clip_y - clip_down, clip_x + clip_left:w - clip_x - clip_right]
        full_detected_table_with_labels = table 
        # cv2.imwrite('table_original_image.jpg', table_original_image)


    ## Detecting the vertical and horizontal (both dotted and bold) in the table using ML algorithms
    # Thresholding to reduce the image to black or white pixels
    if table is None or table.size == 0:
        print(f"Error: A table is not dectected for station: {station}, file: {month_filename}")
        return None  # Exit the function and return None
    else:
        table_img_bin = table

    # # Apply noise reduction
    # table_img_bin = reduce_noise(table_img_bin)

    # Remove only long vertical lines while keeping text intact
    table_img_bin = remove_vertical_lines(table_img_bin)



    # Save the binary image for use later in detecting text
    save_dir = os.path.join(transient_transcription_output_dir, station)
    os.makedirs(save_dir, exist_ok=True)  # Ensure the directory exists
    save_path = os.path.join(save_dir, 'table_binarized.jpg')
    cv2.imwrite(save_path, table_img_bin)
    
    # Invert the binarized image of the table
    img_bin = 255-table_img_bin
    # Detect the vertical lines in the image
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, np.array(table).shape[1]//50)) # The '//50' divides the length of the array (table) by 50, likely to obtain a fraction of the length for the structuring element,
    eroded_image = cv2.erode(img_bin, vertical_kernel, iterations=1)
    vertical_lines = cv2.dilate(eroded_image, vertical_kernel, iterations=5)
    # Detect the horizontal lines in the image
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(table).shape[1]//60, 1)) # The '//20' divides the width of the array (table) by 20, likely to obtain a fraction of the width for the structuring element.
    eroded_image= cv2.erode(img_bin, hor_kernel, iterations=1)
    horizontal_lines = cv2.dilate(eroded_image, hor_kernel, iterations=5)
    # Blending the imaegs with the vertical lines and the horizontal lines 
    combined_vertical_and_horizontal_lines = cv2.addWeighted(vertical_lines, 1, horizontal_lines, 1, 1)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    combined_image_dilated = cv2.dilate(combined_vertical_and_horizontal_lines, kernel, iterations=5)
    # Remove the lines from the image (table)
    image_without_lines = cv2.subtract(img_bin, combined_image_dilated)
    # Remove smaller 'still-visible' lines through noise removal
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    image_without_lines_noise_removed = cv2.erode(image_without_lines, kernel, iterations=1)
    image_without_lines_noise_removed = cv2.dilate(image_without_lines_noise_removed, kernel, iterations=1)
    
    # Detecting the dotted lines using horizontal line detection and erosion. ### ADDITIONAL STEP: This is because the original images ahve dotted horizontal lines which cvan still be detected after the first removal of main (undotted) horizontal lines
    # hor_kernel_2 = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(image_without_lines_noise_removed).shape[1]//100, 1))
    # image_3 = cv2.erode(image_without_lines_noise_removed, hor_kernel_2, iterations=1)
    # horizontal_lines_2 = cv2.dilate(image_3, hor_kernel_2, iterations=5)
    # # Removing the dotted liens by substracting them from the original image 
    # image_without_lines_2 = cv2.subtract(image_without_lines_noise_removed, horizontal_lines_2)

    # #**Removing Remaining Isolated Small Dots (Noise)**
    # dot_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (4,4))  # Detects small noise blobs
    # image_no_dots = cv2.morphologyEx(image_without_lines_noise_removed, cv2.MORPH_OPEN, dot_kernel, iterations=1)

    # plt.imshow(image_no_dots, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.show() 










    # # Step 3: **Using Connected Component Filtering to Remove Tiny Blobs**
    # num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(image_without_lines_noise_removed, connectivity=8)

    # # Create a mask to keep only large components (text)
    # filtered_image = np.zeros_like(image_without_lines_noise_removed)
    # # for i in range(1, num_labels):  # Ignore background (label 0)
    # #     if stats[i, cv2.CC_STAT_AREA] > 100:  # Adjust threshold to remove small dots/noise. 50 was good
    # #         filtered_image[labels == i] = 255  # Keep large components (text)

    # # Set aspect ratio threshold for horizontal noise (adjust as needed)
    # ASPECT_RATIO_THRESHOLD = 4  # If Width / Height > 5, it's considered a horizontal dot chain
    # HEIGHT_THRESHOLD = 28  # Remove any blobs with height less than 28 pixels

    # for i in range(1, num_labels):  # Ignore background (label 0)
    #     x, y, w, h, area = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT], stats[i, cv2.CC_STAT_AREA]
        
    #     aspect_ratio = w / h  # Compute aspect ratio

    #     # Keep components that are NOT horizontal noise
    #     if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and h > HEIGHT_THRESHOLD:  
    #         filtered_image[labels == i] = 255  # Keep text

    # plt.imshow(filtered_image, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.show() 



    # # Step 0: Light dilation to preserve thin strokes
    # kernel_preserve_digits = np.ones((2,2), np.uint8)  
    # image_without_lines_noise_removed = cv2.dilate(image_without_lines_noise_removed, kernel_preserve_digits, iterations=1)







    # # Step 1: Detect all connected components
    # num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(image_without_lines_noise_removed, connectivity=8)

    # # Create a mask to keep only valid text components
    # filtered_image = np.zeros_like(image_without_lines_noise_removed)

    # # Set thresholds
    # ASPECT_RATIO_THRESHOLD = 4  # If Width / Height > 4, it's considered horizontal noise
    # HEIGHT_THRESHOLD = 5  # Remove any blobs with height less than this. was 28 in inital test
    # PROXIMITY_THRESHOLD = 5  # Maximum distance (in pixels) to consider a dot "close" to a number

    # # Step 2: Identify large text components (potential numbers)
    # text_components = []
    # for i in range(1, num_labels):
    #     x, y, w, h, area = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT], stats[i, cv2.CC_STAT_AREA]
        
    #     aspect_ratio = w / h  # Compute aspect ratio
        
    #     # Keep only real text (numbers) as reference components
    #     if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and h > HEIGHT_THRESHOLD:
    #         text_components.append((x, y, w, h))
    #         filtered_image[labels == i] = 255  # Keep text

    # # Step 3: Process small dots and keep only those near text
    # for i in range(1, num_labels):
    #     x, y, w, h, area = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT], stats[i, cv2.CC_STAT_AREA]
        
    #     # Skip already kept text components
    #     if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and h > HEIGHT_THRESHOLD:
    #         continue
        
    #     # Check if this small dot is close to any number
    #     keep_dot = False
    #     for tx, ty, tw, th in text_components:
    #         # Compute distance from dot center to the number bounding box
    #         dot_center_x, dot_center_y = x + w // 2, y + h // 2
    #         if (tx - PROXIMITY_THRESHOLD <= dot_center_x <= tx + tw + PROXIMITY_THRESHOLD and
    #             ty - PROXIMITY_THRESHOLD <= dot_center_y <= ty + th + PROXIMITY_THRESHOLD):
    #             keep_dot = True
    #             break  # No need to check other numbers

    #     if keep_dot:
    #         filtered_image[labels == i] = 255  # Keep dots near numbers

    
    # plt.imshow(filtered_image, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.title('filtered image- no dots - no lines')
    # plt.show() 






    # Step 1: Detect all connected components
    num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(image_without_lines_noise_removed, connectivity=8)

    # Create a mask to keep only valid text components
    filtered_image = np.zeros_like(image_without_lines_noise_removed)

    # Set thresholds
    ASPECT_RATIO_THRESHOLD = 4  # If Width / Height > 4, it's considered horizontal noise
    HEIGHT_THRESHOLD = 5  # Remove any blobs with height less than this
    PROXIMITY_THRESHOLD = 5  # Maximum distance (in pixels) to consider a dot "close" to a number

    # Store y-coordinates and x-limits of horizontal dots **ONLY NEAR TEXT**
    horizontal_lines = []

    # Step 2: Identify large text components (potential numbers)
    text_components = []
    for i in range(1, num_labels):
        x, y, w, h, area = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT], stats[i, cv2.CC_STAT_AREA]

        aspect_ratio = w / h  # Compute aspect ratio

        # Keep only real text (numbers) as reference components
        if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and h > HEIGHT_THRESHOLD:
            text_components.append((x, y, w, h))
            filtered_image[labels == i] = 255  # Keep text

    plt.imshow(filtered_image, cmap="gray")
    plt.title("First Filtered Image - No Dots - No Lines")
    plt.show()

    # Step 3: Process small dots and keep only those near text
    for i in range(1, num_labels):
        x, y, w, h, area = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT], stats[i, cv2.CC_STAT_AREA]

        # Skip already kept text components
        if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and h > HEIGHT_THRESHOLD:
            continue  # Skip large text components

        # Check if this small dot is close to any number
        keep_dot = False
        for tx, ty, tw, th in text_components:
            # Compute distance from dot center to the number bounding box
            dot_center_x, dot_center_y = x + w // 2, y + h // 2
            if (tx - PROXIMITY_THRESHOLD <= dot_center_x <= tx + tw + PROXIMITY_THRESHOLD and
                ty - PROXIMITY_THRESHOLD <= dot_center_y <= ty + th + PROXIMITY_THRESHOLD):
                keep_dot = True
                # Store the y-position and x-limits of dots near text
                horizontal_lines.append((dot_center_y, x, x + w))  # (y-position, x_start, x_end)
                break  # No need to check other numbers

        if keep_dot:
            filtered_image[labels == i] = 255  # Keep dots near numbers

    plt.imshow(filtered_image, cmap="gray")
    plt.title("Second Filtered Image - Only Dots Near Text - No Lines")
    plt.show()

    # **Step 4: Draw Horizontal Separation Lines Only Where Needed**
    #image_with_lines = filtered_image.copy()

    # for y, x_start, x_end in horizontal_lines:
    #     cv2.line(filtered_image, (x_start, y), (x_end, y), 0, thickness=2)  # Use short lines only between detected dots

    # plt.imshow(filtered_image, cmap="gray")
    # plt.title("Image with Artificial Short Horizontal Lines for Row Separation")
    # plt.show()

    

    # Step 2: Use **horizontal dilation** to merge digits within numbers, but prevent full merging
    kernel_to_remove_gaps = np.ones((1, 6), np.uint8)  # Horizontal merging kernel
    image_with_number_blobs = cv2.dilate(filtered_image, kernel_to_remove_gaps, iterations=5)  # Controlled dilation

    # Step 3: Apply **morphological closing** to ensure numbers remain compact blobs
    rect_kernel = np.ones((1, 3), np.uint8)   # was 3 initially. in prev step was 1
    image_with_word_blobs = cv2.morphologyEx(image_with_number_blobs, cv2.MORPH_CLOSE, rect_kernel, iterations=1)


    plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    plt.title('horizontal dilation')
    plt.show() 

    # # Detecting the dotted lines using horizontal line detection and erosion. ### ADDITIONAL STEP: This is because the original images ahve dotted horizontal lines which cvan still be detected after the first removal of main (undotted) horizontal lines
    # hor_kernel_2 = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(image_with_word_blobs).shape[1]//20, 1))   # was 20 originally. 
    # image_3 = cv2.erode(image_with_word_blobs, hor_kernel_2, iterations=1) # had it at 2 originally
    # horizontal_lines_2 = cv2.dilate(image_3, hor_kernel_2, iterations=1)
    # # Removing the dotted liens by substracting them from the original image 
    # image_with_word_blobs = cv2.subtract(image_with_word_blobs, horizontal_lines_2)

    # plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.show() 


    # Define a horizontal erosion kernel
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 6))  # (height, width) to break horizontal merges. had it at 1,5 initially 
    # Apply erosion to break the connection between digits
    image_with_word_blobs = cv2.erode(image_with_word_blobs, horizontal_kernel, iterations=2)  # Increase iterations if still connected. Had 2 iteratons initially
    plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    plt.title('applied horizontal_kernel')
    plt.show() 


    ##************REMOVE THIS***************
    # # Define a vertical erosion kernel
    # vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 1))  # (width, height) to break horizontal merges. had it at 1,5 initially 
    # # Apply erosion to break the connection between digits
    # image_with_word_blobs = cv2.erode(image_with_word_blobs, vertical_kernel, iterations=2)  # Increase iterations if still connected. Had 2 iteratons initially
    # plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.title('applied vertical kernel')
    # plt.show() 
    ##************REMOVE THIS***************



    ##************ PROBABLY REMOVE THIS***************
    # # OPTIONAL: Apply light dilation to restore digit thickness
    # image_with_word_blobs = cv2.dilate(image_with_word_blobs, np.ones((3, 1), np.uint8), iterations=2) # (height, width)
    # plt.imshow(image_with_word_blobs, cmap = 'gray') 
    # plt.title('final vertical dilation to increase height of blobs')
    # plt.show() 
    ##************REMOVE THIS***************







    # # Find connected components
    # num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(image_with_word_blobs, connectivity=8)

    # # Create a mask for filtered output
    # filtered_image = image_with_word_blobs.copy()

    # # Define a vertical erosion kernel
    # vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (7, 1))  # (height, width) to break vertical merges

    # # **Loop Through Each Component and Apply More Erosion to Taller Blobs**
    # for i in range(1, num_labels):  # Ignore background (label 0)
    #     x, y, w, h = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT]

    #     # Check if blob is taller than expected
    #     if h > max_cell_height_threshold:
    #         erosion_iterations = 5  # Apply stronger vertical erosion to break stacked numbers
    #     else:
    #         erosion_iterations = 1

    #     # Extract the region of interest (ROI) for erosion
    #     roi = image_with_word_blobs[y:y+h, x:x+w]

    #     # Apply vertical erosion only on the specific region
    #     eroded_roi = cv2.erode(roi, vertical_kernel, iterations=erosion_iterations)

    #     # Put the processed region back into the filtered image
    #     filtered_image[y:y+h, x:x+w] = eroded_roi


    # image_with_word_blobs = filtered_image
    # plt.imshow(image_with_word_blobs, cmap = 'gray') 
    # plt.title('all large vertical blobs joining different rows removed')
    # plt.show() 



    # Create a blank mask for horizontal lines
    horizontal_line_mask = np.zeros_like(image_with_word_blobs)

    # Draw **thin** horizontal lines at detected y-positions (but not subtracting yet)
    for y, x_start, x_end in horizontal_lines:
        cv2.line(horizontal_line_mask, (x_start - 15, y), (x_end + 15, y), 255, thickness=3)  # Initially very thin

    # Show the initial thin line mask
    plt.imshow(horizontal_line_mask, cmap="gray")
    plt.title("Initial Thin Line Mask")
    plt.show()

    # # **Step 2: Dilate the horizontal lines to make them stronger**
    # dilation_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 2))  # Wider lines
    # horizontal_line_mask = cv2.dilate(horizontal_line_mask, dilation_kernel, iterations=1)

    # # Show the dilated horizontal line mask
    # plt.imshow(horizontal_line_mask, cmap="gray")
    # plt.title("Dilated Horizontal Line Mask")
    # plt.show()

    # **Step 3: Subtract the dilated lines from the filtered image**
    image_with_word_blobs = cv2.subtract(image_with_word_blobs, horizontal_line_mask)

    # Show the final image after subtraction
    plt.imshow(image_with_word_blobs, cmap="gray")
    plt.title("Final Image After Subtracting Horizontal Lines")
    plt.show()



    ####*********what if i did this step above now with (1,3) dilation. In otherwords in both direction
    image_with_word_blobs = cv2.dilate(image_with_word_blobs, np.ones((1, 6), np.uint8), iterations=2) # (height, width) # initialy (1,3) and 1 iteration
    plt.imshow(image_with_word_blobs, cmap = 'gray') 
    plt.title('final horizontal dilation to increase width of blobs')
    plt.show() 

    # **Step 2: Apply Controlled Erosion to Fix Over-Merging**
    erosion_kernel = np.ones((1, 3), np.uint8)  # Smaller horizontal erosion kernel
    image_with_word_blobs = cv2.erode(image_with_word_blobs, erosion_kernel, iterations=1)  # Controlled shrinking

    plt.imshow(image_with_word_blobs, cmap="gray")
    plt.title("After Horizontal Erosion (Fix Over-Merging)")
    plt.show()

    # **Step 2: Find Wide Blobs That Need Stronger Erosion**
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(image_with_word_blobs, connectivity=8)

    for i in range(1, num_labels):  # Ignore background (label 0)
        x, y, w, h = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT]
        
        if w > 200:  # **If the blob is too wide, apply stronger erosion**
            erosion_kernel = np.ones((3, 10), np.uint8)  # Slightly stronger erosion for wide blobs
            roi = image_with_word_blobs[y:y+h, x:x+w]
            eroded_roi = cv2.erode(roi, erosion_kernel, iterations=3)  # More iterations for wider blobs'
            
            # **Apply vertical recovery dilation to restore eroded thickness**
            dilation_kernel = np.ones((2, 1), np.uint8)  # Small dilation to restore height
            recovered_roi = cv2.dilate(eroded_roi, dilation_kernel, iterations=3)  # Restore vertical thickness
            
            image_with_word_blobs[y:y+h, x:x+w] = recovered_roi  # Replace only this region
        
        if h > 50: # ** If the blob is too tall, apply stronger erosion along the horizontal dimension**
            # erosion_kernel_vert = np.ones((min(h // 15, 12), 6), np.uint8)  # Adaptive vertical erosion # Initially had this: np.ones((min(h // 20, 12), 3), np.uint8)
            erosion_kernel_vert = np.ones((1, 13), np.uint8)  # Slightly stronger erosion for wide blobs # Imitially had (10,3), BUT (1,10) improved it
            roi = image_with_word_blobs[y:y+h, x:x+w]
            eroded_roi = cv2.erode(roi, erosion_kernel_vert, iterations=3)  
            
            # **Restore horizontal thickness**
            dilation_kernel_vert = np.ones((1, 6), np.uint8)  # initially had (1,6)
            recovered_roi = cv2.dilate(eroded_roi, dilation_kernel_vert, iterations=3)  

            image_with_word_blobs[y:y+h, x:x+w] = recovered_roi 

    
    plt.imshow(image_with_word_blobs, cmap="gray")
    plt.title("After Adaptive Erosion (Fixing Over-Merging)")
    plt.show()

    
    
    
    
    
    
 






# ****TO BE DELETED. METHOD DIDNT WORK

#    image_with_word_blobs = cv2.dilate(image_with_word_blobs, np.ones((2, 1), np.uint8), iterations=2) # (height, width) 
#     plt.imshow(image_with_word_blobs, cmap = 'gray') 
#     plt.title('final vertical dilation to increase width of blobs')
#     plt.show() 


#     # **Step 3: Subtract the Vertical Lines from the Image to Break Column Merges**
#    # **Step 1: Preprocess to Enhance Gaps (Optional)**
#     image_with_word_blobs = cv2.morphologyEx(image_with_word_blobs, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))

#     # **Step 2: Compute the Vertical Projection Profile**
#     vertical_projection = np.sum(255 - image_with_word_blobs, axis=0)  # Sum pixel intensities along each column

#     # **Step 3: Normalize & Identify Vertical Gaps**
#     normalized_projection = vertical_projection / (np.max(vertical_projection) + 1e-5)  # Avoid division by zero
#     threshold = 0.2  # Adjust this to detect prominent dark vertical gaps
#     column_positions = np.where(normalized_projection < threshold)[0]  # Get columns with the lowest intensity

#     # **Step 4: Filter Column Positions for Clean Separation**
#     filtered_columns = []
#     min_distance = 50  # Avoid closely spaced columns

#     for i in range(len(column_positions) - 1):
#         if (i == 0) or (column_positions[i] - column_positions[i - 1] > min_distance):
#             filtered_columns.append(column_positions[i])

#     # **Step 5: Draw Vertical Lines Using These Positions**
#     overlay_image = cv2.cvtColor(image_with_word_blobs, cv2.COLOR_GRAY2BGR)  # Convert grayscale to color for visualization

#     for x in filtered_columns:
#         cv2.line(overlay_image, (x, 0), (x, image_with_word_blobs.shape[0]), (0, 0, 255), thickness=3)  # Red lines

#     # **Step 6: Show the Overlayed Image Before Subtraction**
#     plt.figure(figsize=(10, 6))
#     plt.imshow(overlay_image)
#     plt.title("Image with Automatically Detected Vertical Lines")
#     plt.show()

#     # **Step 6: Create a Mask for Subtraction**
#     vertical_line_mask = np.zeros_like(image_with_word_blobs)
#     for x in filtered_columns:
#         cv2.line(vertical_line_mask, (x, 0), (x, image_with_word_blobs.shape[0]), 255, thickness=3)  # Thicker mask lines

#     # **Step 7: Subtract Vertical Lines from the Image**
#     image_with_word_blobs = cv2.subtract(image_with_word_blobs, vertical_line_mask)

#     # **Step 8: Show Final Image After Removing Vertical Connections**
#     plt.imshow(image_with_word_blobs, cmap="gray")
#     plt.title("Final Image After Removing Vertical Merges")
#     plt.show()






    # kernel_to_remove_gaps_between_words = np.ones((1, 10), np.uint8)  # Larger kernel to bridge gaps better
    # image_with_word_blobs = cv2.dilate(image_with_word_blobs, kernel_to_remove_gaps_between_words, iterations=1) # was 5 iterations previously

    # # kernel_to_remove_gaps_between_words = np.ones((1, 23), np.uint8)  # Larger kernel to bridge gaps better
    # # image_with_word_blobs = cv2.dilate(image_with_word_blobs, kernel_to_remove_gaps_between_words, iterations=1) # was 5 iterations previously

    # plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.show() 





    # # simple_kernel = np.ones((3,3), np.uint8)
    # simple_kernel = np.ones((1,3), np.uint8)  # 1 row, 3 columns
    # image_with_word_blobs = cv2.dilate(image_with_word_blobs, simple_kernel, iterations=2)

    # # === NEW STEP: Detect & Remove Short Horizontal Lines (Above 100px) ===
    # min_line_length = 50  # Threshold for what counts as a horizontal line
    # max_line_gap = 10  # Allow some small breaks in detected lines

    # # Apply Canny edge detection to find edges
    # edges = cv2.Canny(image_with_word_blobs, 50, 150, apertureSize=7)

    # # Detect lines using Hough Transform
    # lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=50, minLineLength=min_line_length, maxLineGap=max_line_gap)

    # # Remove detected horizontal lines
    # if lines is not None:
    #     for line in lines:
    #         x1, y1, x2, y2 = line[0]
    #         angle = np.abs(np.arctan2(y2 - y1, x2 - x1) * 180.0 / np.pi)

    #          # Remove lines if they are roughly horizontal (angle between -15° to +15°)
    #         if (angle < 15) or (angle > 165):  
    #             cv2.line(image_with_word_blobs, (x1, y1), (x2, y2), (0, 0, 0), thickness=2)  # Overwrite with black



    # # Detecting the dotted lines using horizontal line detection and erosion. ### ADDITIONAL STEP: This is because the original images ahve dotted horizontal lines which cvan still be detected after the first removal of main (undotted) horizontal lines
    # hor_kernel_2 = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(image_with_word_blobs).shape[1]//55, 1))
    # image_3 = cv2.erode(image_with_word_blobs, hor_kernel_2, iterations=2)
    # horizontal_lines_2 = cv2.dilate(image_3, hor_kernel_2, iterations=1)
    # # Removing the dotted liens by substracting them from the original image 
    # image_without_lines_2 = cv2.subtract(image_with_word_blobs, horizontal_lines_2)


    ## Using contours in order to detect text in the table after removing the vertical and horizontal lines
    # Assuming 'table' is your input image in BGR format
    # result = cv2.findContours(image_without_lines_2, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    result = cv2.findContours(image_with_word_blobs, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contours = result[0]
    # Original image of table in binarizesd format
    image_with_all_bounding_boxes = cv2.imread(save_path)
    table_binarized = image_with_all_bounding_boxes.copy()
    

    # Plots only for visualization purposes. Uncomment the lines below to show the different steps
    
    plt.imshow(image_without_lines, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    plt.show() 

    plt.imshow(image_without_lines_noise_removed, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    plt.show() 

    plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing text blobs on the detected table image with horizintal and vertical lines removed.
    plt.show()

    # plt.imshow(image_without_lines_2, cmap = 'gray') # figure showing text blobs on the detected table image with horizintal and vertical lines removed.
    # plt.show()

    # plt.imshow(detected_table_cells[4], cmap = 'gray') # unclipped detected table
    # plt.show()

    # plt.imshow(detected_table_cells[1]) # clipped detected table
    # plt.show()

    ## FOR VISUALIZATION PURPOSES. Uncomment the lines below to plot the identified cells (contours/bounding boxes)
    # Make a copy of the original image to overlay contours without modifying the original
    table_img_bin_overlayed_with_contours = table_img_bin.copy()
    # Convert the grayscale image to RGB to support colored bounding boxes
    table_img_bin_overlayed_with_contours = cv2.cvtColor(table_img_bin_overlayed_with_contours, cv2.COLOR_GRAY2RGB)


    # Filter out smaller or larger bounding boxes from all the detected text contours. This is helpful to avoid overly large cells or small cells with no text. Remember to adjust these values based on the table structure in your specific case 
    filtered_contours = filter_contours(contours, min_cell_width_threshold, min_cell_height_threshold, max_cell_width_threshold, max_cell_height_threshold)

    # Iterate over each contour in the new_contours list and draw bounding boxes
    for contour in filtered_contours:
        if contour is not None and len(contour) > 0:
            x, y, w, h = cv2.boundingRect(contour)

            # Adjust bounding box dimensions
            increase_factor_width = 0.05
            increase_factor_height = 0.25
            x += int(w * increase_factor_width) # Increase width
            y -= int(h * increase_factor_height) # Increase height
            w -= int(w * increase_factor_width) # Decrease width a little to avoid vertical lines that may be transcribed as the number 1 yet they aren't a number
            h += int(h * increase_factor_height * 2) # Increase height
            
            # Draw the bounding box directly on the overlay image
            cv2.rectangle(table_img_bin_overlayed_with_contours, (x, y), (x + w, y + h), (0, 0, 255), 4)

    # Display the image with bounding boxes using matplotlib
    plt.imshow(table_img_bin_overlayed_with_contours)
    plt.axis('off')  # Hide axis
    plt.show()

    # Sort contours by y-coordinate
    contours_sorted = sorted(filtered_contours, key=lambda c: cv2.boundingRect(c)[1])

    # Get the dimensions of the loaded image. Here, particulary the image/table width is very important for the column placement of cells/bounding boxes
    image_height, image_width, image_channels = image_with_all_bounding_boxes.shape


    # Adding missing bounding boxes. Here, we define the minimum height space and minimum width space between the bounding boxes in a column and row respectively, in case of a missing bounding box.
    # Add missing ROIs to the contours
    new_contours = add_missing_rois(contours_sorted, space_height_threshold, space_width_threshold, max_cell_height_per_box, no_of_rows, no_of_columns, image_width)
    

    ## FOR VISUALIZATION PURPOSES. Uncomment the lines below to plot the identified cells (contours/bounding boxes)
    # Make a copy of the original image to overlay contours without modifying the original
    table_img_bin_overlayed_with_contours = table_img_bin.copy()
    # Convert the grayscale image to RGB to support colored bounding boxes
    table_img_bin_overlayed_with_contours = cv2.cvtColor(table_img_bin_overlayed_with_contours, cv2.COLOR_GRAY2RGB)


    # Iterate over each contour in the new_contours list and draw bounding boxes
    for contour in new_contours:
        if contour is not None and len(contour) > 0:
            x, y, w, h = cv2.boundingRect(contour)

            # Adjust bounding box dimensions
            increase_factor_width = 0.05
            increase_factor_height = 0.25
            x += int(w * increase_factor_width) # Increase width
            y -= int(h * increase_factor_height) # Increase height
            w -= int(w * increase_factor_width) # Decrease width a little to avoid vertical lines that may be transcribed as the number 1 yet they aren't a number
            h += int(h * increase_factor_height * 2) # Increase height
            
            # Draw the bounding box directly on the overlay image
            cv2.rectangle(table_img_bin_overlayed_with_contours, (x, y), (x + w, y + h), (0, 255, 0), 5)

    # Display the image with bounding boxes using matplotlib
    plt.imshow(table_img_bin_overlayed_with_contours)
    plt.axis('off')  # Hide axis
    plt.show()


    detected_table_cells = [new_contours, image_with_all_bounding_boxes, table_binarized, table_original_image, full_detected_table_with_labels]
    

    return detected_table_cells

