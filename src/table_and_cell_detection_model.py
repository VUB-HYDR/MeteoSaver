#!/usr/bin/env python
import os, argparse, glob, tempfile, shutil, warnings
import cv2
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans

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
    Rotates the entire image (page) while keeping text naturally aligned.

    This function detects the skew angle of horizontal lines and rotates the page accordingly.
    The text remains visually undistorted because it rotates with the page, maintaining correct orientation.

    Parameters
    --------------
    image : 
        The input image that needs to be deskewed.

    Returns
    --------------
    deskewed_image : 
        The deskewed image after rotation.
    '''

    # Detect horizontal lines and calculate the average skew angle
    hor_contours = detect_lines(image, (np.array(image).shape[1] // 20, 1), iterations=1)
    hor_angle = calculate_average_angle(hor_contours, orientation='horizontal')

    # If the detected angle is too small, no need to rotate
    if abs(hor_angle) < 0.1:
        print("[DEBUG] No significant skew detected. Returning original image ...")
        return image

    # Get image dimensions
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)  # Compute the center of rotation

    # Compute rotation matrix
    M = cv2.getRotationMatrix2D(center, -hor_angle, 1.0)  # Rotate page to align it

    # Rotate the image using affine transformation
    deskewed_image = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    print(f"[DEBUG] Detected skew angle: {hor_angle:.2f} degrees (Page Rotated) ...")

    return deskewed_image, hor_angle



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



def add_missing_rois(sorted_contours, space_threshold, space_width_threshold, max_cell_height_per_box, max_rows, num_columns, image_width):
    '''
    Improved function to add missing ROIs **only where they align with existing rows in at least two neighboring columns**.
    '''

    # ✅ Step 1: Group contours into columns
    columns = group_contours_into_columns(sorted_contours, num_columns, image_width)

    # ✅ Collect all row centers across all columns
    all_row_centers = []
    for i in sorted(columns.keys()):
        column_boxes = sorted(columns[i], key=lambda b: b[1])  # Sort by y-coordinate
        all_row_centers.extend([box[1] + box[3] // 2 for box in column_boxes])

    # ✅ Step 2: Apply KMeans for row clustering
    if len(all_row_centers) > 1:
        num_clusters = min(max_rows, len(all_row_centers))  # Ensure at least 1 cluster
        kmeans = KMeans(n_clusters=num_clusters, random_state=42, n_init=10)
        kmeans.fit(np.array(all_row_centers).reshape(-1, 1))
        global_row_centroids = sorted(kmeans.cluster_centers_.flatten())
    else:
        global_row_centroids = np.linspace(0, max(all_row_centers), max_rows)  # Default spacing

    print(f'[DEBUG] Global estimated row centroids = {global_row_centroids}')

    new_boxes = []

    # ✅ Step 3: Process each column
    for i in sorted(columns.keys()):  # Process columns in order
        column_boxes = sorted(columns[i], key=lambda b: b[1])  # Sort by y-coordinate
        column_count = len(column_boxes)
        print(f'[DEBUG] Column {i}: {column_count} detected rows.')

        # ✅ Identify gaps between detected cells
        gaps = []
        for j in range(1, len(column_boxes)):
            prev_box = column_boxes[j - 1]
            curr_box = column_boxes[j]
            space_between = curr_box[1] - (prev_box[1] + prev_box[3])

            if space_between > space_threshold:
                gaps.append((space_between, prev_box, curr_box))

        # ✅ Sort gaps from largest to smallest
        gaps.sort(key=lambda x: x[0], reverse=True)

        # ✅ Step 4: Add missing boxes **only where they align with neighboring columns**
        missing_count = max_rows - column_count
        for gap in gaps:
            if column_count >= max_rows or missing_count <= 0:
                break  # Stop if we exceed max_rows

            space_between, prev_box, curr_box = gap
            new_y = prev_box[1] + prev_box[3] + (space_between - max_cell_height_per_box) // 2

            # Adaptive expansion of the box width
            expansion_factor = 1.2  # Increase width by 20% to ensure text coverage
            new_x = prev_box[0] - ((expansion_factor - 1) * space_width_threshold) // 2
            new_width = int(space_width_threshold * expansion_factor)
            new_box = (new_x, new_y, new_width, max_cell_height_per_box)

            # ✅ Find how many existing boxes in **neighboring columns** align with `new_y`
            aligned_boxes = []
            if i > 0:  # Check previous column
                aligned_boxes += [box for box in columns.get(i - 1, []) if abs((box[1] + box[3] // 2) - (new_y + max_cell_height_per_box // 2)) < 15]
            if i < num_columns - 1:  # Check next column
                aligned_boxes += [box for box in columns.get(i + 1, []) if abs((box[1] + box[3] // 2) - (new_y + max_cell_height_per_box // 2)) < 15]

            # ✅ Make sure the new box aligns with at least **two** neighbors
            if len(aligned_boxes) >= 2:
                print(f'[DEBUG] Adding missing box at: {new_box} (aligned with 2+ neighboring columns)')
                column_boxes.append(new_box)
                column_count += 1
                missing_count -= 1
            elif any(abs(centroid - (new_y + max_cell_height_per_box // 2)) < 20 for centroid in global_row_centroids):  # tried with 15 initially
                print(f'[DEBUG] Adding missing box at: {new_box} (aligned with row centroid)')
                column_boxes.append(new_box)
                column_count += 1
                missing_count -= 1

        # ✅ Re-sort column after adding missing boxes
        column_boxes = sorted(column_boxes, key=lambda b: b[1])
        new_boxes.extend(column_boxes)

    # ✅ Convert bounding boxes into OpenCV contours
    new_contours = [np.array([
        [box[0], box[1]], [box[0] + box[2], box[1]], 
        [box[0] + box[2], box[1] + box[3]], [box[0], box[1] + box[3]]
    ], dtype=np.int32) for box in new_boxes]

    return new_contours




# --- Noise Reduction ---
def reduce_noise(image):
    """Apply morphological operations to reduce small dots and noise."""
    kernel = np.ones((2,2), np.uint8)  # Adjust the kernel size as needed
    return cv2.morphologyEx(image, cv2.MORPH_OPEN, kernel, iterations=1)



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
    
    ## STEP 1: DETECT THE TABLE FROM AN IMAGE/SHEET
    ## Using the pre-processing image to detect the table from the record sheets
    # Here, the threshold value for pixel intensities = 0, and the value 255 is assigned if the pixel value is above the threshold
    thresh = cv2.threshold(image_in_grayscale,0,255,cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1] 
    # Perform morphological operations (like dilation and erosion) for better segmentation
    kernel = np.ones((5,5),np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # Minimum dimensions of table
    threshold_area = min_cell_height_threshold * min_cell_width_threshold  # Minimum contour area to consider as a cell
    threshold_height = min_cell_height_threshold   # Minimum height of the cell
    threshold_width = min_cell_width_threshold    # Minimum width of the cell
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
    ## Only for visualization purposes. Uncomment this to plot the detected table from the entire image/sheet
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
            
            table, skew_angle = deskew(table) # Deskew the image, # Optional, uncomment if you'd like to use this: Incase some of your images are skewed.
            
            table_original_image = original_image[y + clip_up:y + h - clip_down , x + clip_left:x + w - clip_right] # clip out the table (here, the largest contour) from the original image. ** - 420 here to clip out the header rows from the table image and -270 is for the below the table
            # Apply same rotation to original image
            (h, w) = table_original_image.shape[:2]
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, -skew_angle, 1.0)
            table_original_image = cv2.warpAffine(table_original_image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
            # cv2.imwrite('table_original_image.jpg', table_original_image)

        else: # This indicates that the actual table was not detected from the image rather the whole sheet as a the table (for example, due to thick page boarders detected as a table)
            # We there make use of the knowledge of the average table dimensions (in pixels) in relation to the images that were determined from our sample sheets in the dataset to determine location of the table
            # This is thus a bug fix i.e. the MANUAL alternative to the table detection, where the AUTO-DETECTION does not detect the actual table.
            # Calculate the amount to clip from each side
            clip_x = (width - max_table_width) // 2  # Approximate table width in pixels = 3900. # Adjust these values according to your table 
            clip_y = (height - max_table_height) // 2 # Approximate table height in pixels = 3600. # Adjust these values according to your table 

            table = binarized_image[y + clip_y + 630:y + h - clip_y - 300, x + clip_x + 350:x + w - clip_x - 180]  # Here we manually clip the sheets to ensure clipping of the HEADERS and ROW LABELS (Date & Pentad no. in our case) from the table (table detected manually). Adjust this to your case study.
            table_original_image = original_image[y + clip_y + 630:y + h - clip_y - 300, x + clip_x + 350:x + w - clip_x - 180]
                    
    else:
        # If no largest contour is detected. This indicates that the NO table was not detected from the image. Therefore we use the  entire image and make use of the knowledge of the average table dimensions (in pixels) in relation to the images that were determined from our sample sheets in the dataset to determine location of the table.
        # This is thus the MANUAL alternative to the table detection.
        height, width = image_in_grayscale.shape[:2]
        x, y, w, h = 0, 0, width, height  # Consider the entire image dimensions

        # Calculate the amount to clip from each side
        clip_x = (width - max_table_width) // 2
        clip_y = (height - max_table_height) // 2

        table = binarized_image[y + clip_y + 630:y + h - clip_y - 300, x + clip_x + 350:x + w - clip_x - 180]  # Here we manually clip the sheets to ensure clipping of the HEADERS and ROW LABELS (Date & Pentad no. in our case) from the table (table detected manually). Adjust this to your case study.
        table_original_image = original_image[y + clip_y + 630:y + h - clip_y - 300, x + clip_x + 350:x + w - clip_x - 180]
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
        table_img_bin = table # detected table as a binary image

    # Save the binary image for use later in detecting text
    save_dir = os.path.join(transient_transcription_output_dir, station)
    os.makedirs(save_dir, exist_ok=True)  # Ensure the directory exists
    save_path = os.path.join(save_dir, 'table_binarized.jpg')
    cv2.imwrite(save_path, table_img_bin)
    


    ## STEP 2: DETECT THE HORIZONTAL AND VERTICAL LINES OF THE TABLE
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
    
    # Blending the images with the vertical lines and the horizontal lines 
    combined_vertical_and_horizontal_lines = cv2.addWeighted(vertical_lines, 1, horizontal_lines, 1, 1)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    combined_image_dilated = cv2.dilate(combined_vertical_and_horizontal_lines, kernel, iterations=5)
    
    # # Only for visualization purposes. Uncomment the lines below to visualize the detected horizontal and vertical lines in the table
    # plt.imshow(combined_image_dilated, cmap="gray")
    # plt.title("combined_lines_image")
    # plt.show()

    # Remove the lines from the image (table)
    image_without_lines = cv2.subtract(img_bin, combined_image_dilated)
    
    # Remove smaller 'still-visible' lines through noise removal
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    image_without_lines_noise_removed = cv2.erode(image_without_lines, kernel, iterations=1)
    image_without_lines_noise_removed = cv2.dilate(image_without_lines_noise_removed, kernel, iterations=1)
    


    ## STEP 3: DETECT AND REMOVE THE NOISE SUCH AS SMALL DOTS ON OUR TABLE IMAGE
    # Detect all connected (text) components
    num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(image_without_lines_noise_removed, connectivity=8)

    # Create a mask to keep only valid text components
    filtered_image = np.zeros_like(image_without_lines_noise_removed)

    # Set text thresholds (in pixels)
    ASPECT_RATIO_THRESHOLD = 4  # If Width / Height > 4 (was 4 before this new check), it's considered horizontal noise
    HEIGHT_THRESHOLD = 5  # Remove any blobs with height less than this
    PROXIMITY_THRESHOLD = 5  # Maximum distance (in pixels) to consider a dot "close" to a number

    # Store y-coordinates and x-limits of horizontal dots **ONLY NEAR TEXT**
    horizontal_lines = []
  
    # Pre-compute aspect ratios for all components
    if num_labels > 1:
        widths = stats[1:, cv2.CC_STAT_WIDTH]
        heights = stats[1:, cv2.CC_STAT_HEIGHT]
        areas = stats[1:, cv2.CC_STAT_AREA]
        aspect_ratios = widths / heights

        text_components = []
        text_labels_to_keep = []
        dot_labels_to_remove = []

        # Single loop with pre-computed values
        for i in range(num_labels - 1):  # Exclude background
            label_id = i + 1
            area = areas[i]
            aspect_ratio = aspect_ratios[i]
            height = heights[i]

            # Keep only real text (numbers) as reference components
            if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and height > HEIGHT_THRESHOLD:
                x, y, w, h = stats[label_id, cv2.CC_STAT_LEFT:cv2.CC_STAT_LEFT+4]
                text_components.append((x, y, w, h))
                text_labels_to_keep.append(label_id)
            else:
                dot_labels_to_remove.append(label_id)

        # Batch update images
        if text_labels_to_keep:
            text_mask = np.isin(labels, text_labels_to_keep)
            filtered_image[text_mask] = 255

        dots_to_remove = np.zeros_like(labels)
        if dot_labels_to_remove:
            dot_mask = np.isin(labels, dot_labels_to_remove)
            dots_to_remove[dot_mask] = 255
    else:
        text_components = []
        dots_to_remove = np.zeros_like(labels)
    
    # Pre-convert text_components to numpy arrays for faster operations
    if text_components:
        text_array = np.array(text_components)  # Shape: (N, 4) for N text components

        for i in range(1, num_labels):
            x, y, w, h, area = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT], stats[i, cv2.CC_STAT_AREA]
            aspect_ratio = w / h
            # Skip already kept text components
            if area > 100 and aspect_ratio < ASPECT_RATIO_THRESHOLD and h > HEIGHT_THRESHOLD:
                continue

            # Vectorized proximity check
            dot_center_x, dot_center_y = x + w // 2, y + h // 2

            # Check proximity to all text components at once
            in_x_range = ((text_array[:, 0] - PROXIMITY_THRESHOLD <= dot_center_x) & 
                        (dot_center_x <= text_array[:, 0] + text_array[:, 2] + PROXIMITY_THRESHOLD))
            in_y_range = ((text_array[:, 1] - PROXIMITY_THRESHOLD <= dot_center_y) & 
                        (dot_center_y <= text_array[:, 1] + text_array[:, 3] + PROXIMITY_THRESHOLD))
            
            keep_dot = np.any(in_x_range & in_y_range)

            if keep_dot:
                filtered_image[labels == i] = 255
                horizontal_lines.append((dot_center_y, x, x + w))

    # # Only for visualization purposes. Uncomment the lines below to visualize the filtered table image showing text/digits without the noise such as dots.
    # plt.imshow(filtered_image, cmap="gray")
    # plt.title("Second Filtered Image - Only Dots Near Text - No Lines")
    # plt.show()
    


    ## STEP 4: DETECT THE ALL CONNECTED TEXT/MULTI-DIGITS IN ONE CELL (text/digit locations)
    # Blobs refer to white connected components (regions) on the **inverted** image.
    # These components likely represent text or digits within a cell — especially when characters are close together.
    # To improve detection of full numbers or words (e.g., merging '1' and '9' in '19' into a single blob), we use **dilation**.

    # Dilation expands white regions in the image. In this context, it helps connect nearby digits/letters horizontally 
    # (e.g., connecting the characters in "27" that may have a small gap between them).
    # The kernel shape (1, 6) means we only merge components that are close **horizontally**, not vertically.
    kernel_to_remove_gaps = np.ones((1, 6), np.uint8)  # Horizontal merging kernel. Variables are (Height, Width) in pixels
    image_with_number_blobs = cv2.dilate(filtered_image, kernel_to_remove_gaps, iterations=5)  # Controlled dilation

    # Next, we apply **morphological closing** to clean up the dilated blobs.
    # Morphological closing = Dilation followed by Erosion. It connects small gaps and fills tiny holes within each blob,
    # making digits appear more solid and compact — especially useful for unifying thin or fragmented characters.
    rect_kernel = np.ones((1, 6), np.uint8)   
    image_with_word_blobs = cv2.morphologyEx(image_with_number_blobs, cv2.MORPH_CLOSE, rect_kernel, iterations=1)

    # # Only for visualization purposes. Uncomment the lines below to visualize the table image showing blobs (text/digit locations)
    # plt.imshow(image_with_word_blobs, cmap = 'gray') 
    # plt.title('horizontal dilation and morpholodical closing')
    # plt.show() 

    # After this, we subtract the detected vertical and horizontal lines from the filtered image.
    # This is to prevent over-merging of digits/text across different columns or rows
    image_with_word_blobs = cv2.subtract(image_with_word_blobs, combined_image_dilated)

    # # Only for visualization purposes. Uncomment the lines below to visualize the table image showing blobs (text/digit locations), after deleting the original horizontal and vertical lines in the table to avoid over-merging of blobs incorrectly across columns and rows
    # plt.imshow(image_with_word_blobs, cmap = 'gray') # figure showing detected table image with horizintal and vertical lines removed.
    # plt.title('subtraction of detected vertical and horizontal lines of table')
    # plt.show() 

    ## RECHECK: AVOIDING WRONGLY CONNECTED BLOBS/CELLS ACROSS THE TABLE
    ## Identifying wrongfully joined blobs/cells in the same column and adjascent rows
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(image_with_word_blobs, connectivity=8)

    for i in range(1, num_labels):  # Ignore background (label 0)
        x, y, w, h = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT]
        
        # Here we check the heights of all the detected blobs individually
        # If the blob is too tall, we slice it to the expected maximum dimensions, here the maximum cell height, to separate wrongfully joined blobs in the same column and adjascent rows
        if h > max_cell_height_threshold: 
            roi = image_with_word_blobs[y:y+h, x:x+w].copy()
            sliced_roi = roi.copy()

            # Determine how many rows this blob spans
            row_height = max_cell_height_threshold  # Approximate average row height in the sheets
            num_slices = h // row_height

            # Insert horizontal black lines to slice between rows to slice the blobs
            for s in range(1, num_slices + 1):
                y_line = s * row_height
                if y_line < h:
                    cv2.line(sliced_roi, (0, y_line), (w, y_line), 0, thickness=3)
                    # print('Slided one blob horizontally') # Just a check to know how many wrongfully joined blobs in the same column and adjascent rows were sliced. Uncomment this incase you want to print the text everytime a blob is sliced.

            # Replace original ROI with sliced version
            image_with_word_blobs[y:y+h, x:x+w] = sliced_roi

    # # For visualization purposes; to ensure that all previously over-merged cells in the same column have been separated by slicing
    # plt.imshow(image_with_word_blobs, cmap="gray")
    # plt.title("Fixing Over-Merging across rows")
    # plt.show()
    
    # Identifying wrongfully joined blobs/cells in the same row and adjascent columns
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(image_with_word_blobs, connectivity=8)

    for i in range(1, num_labels):  # Ignore background (label 0)
        x, y, w, h = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT]

        # Here we check the widths of all the detected blobs individually
        # If the blob is too wide, apply stronger erosion to separate the wrongfully joined blobs in the same row and adjascent columns 
        if w > max_cell_width_threshold:  
            erosion_kernel = np.ones((3, 1), np.uint8)  # where (Height, Width).  Slightly stronger erosion for wide blobs
            roi = image_with_word_blobs[y:y+h, x:x+w]
            eroded_roi = cv2.erode(roi, erosion_kernel, iterations=3)  # More iterations for wider blobs'
            
            # **Apply vertical recovery dilation to restore eroded thickness**
            dilation_kernel = np.ones((2, 1), np.uint8)  # Small dilation to restore height
            recovered_roi = cv2.dilate(eroded_roi, dilation_kernel, iterations=3)  # Restore vertical thickness
            
            image_with_word_blobs[y:y+h, x:x+w] = recovered_roi  # Replace only this region

            # print('eroded one blob') # Just a check to know how many wrongfully joined blobs in the same row and adjascent columns were sliced. Uncomment this incase you want to print the text everytime a blob is sliced.

    # Following up on the previous identification of wrongfully joined blobs/cells in the same row and adjascent columns. Here we find very wide blobs/cells where the erosion (above) wasn't strong enough to separate them, then we slice themm using the maximum cell width threshold**
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(image_with_word_blobs, connectivity=8)

    for i in range(1, num_labels):  # Ignore background (label 0)
        x, y, w, h = stats[i, cv2.CC_STAT_LEFT], stats[i, cv2.CC_STAT_TOP], stats[i, cv2.CC_STAT_WIDTH], stats[i, cv2.CC_STAT_HEIGHT]

        if w > max_cell_width_threshold:
            roi = image_with_word_blobs[y:y+h, x:x+w].copy()
            sliced_roi = roi.copy()

            # Determine how many columns this blob spans
            col_width = max_cell_width_threshold  # Approximate average column width in your dataset
            num_slices = w // col_width

            # Insert vertical black lines to slice between columns
            for s in range(1, num_slices + 1):
                x_line = s * col_width
                if x_line < w:
                    cv2.line(sliced_roi, (x_line, 0), (x_line, h), 0, thickness=3)
                    # print('Sliced one blob vertically')  # For debugging

            # Replace original ROI with sliced version
            image_with_word_blobs[y:y+h, x:x+w] = sliced_roi
            # print('sliced one blob that couldnt be eroded easily') # Just a check to know how many wrongfully joined blobs in the same row and adjascent columns were sliced. Uncomment this incase you want to print the text everytime a blob is sliced.
    
    # # For visualization purposes; to ensure that all previously over-merged cells in same row have been separated by erosion and slicing
    # plt.imshow(image_with_word_blobs, cmap="gray")
    # plt.title("Fixing Over-Merging across columns")
    # plt.show()

    # RECHECK: Remove the table vertical and horizontal lines from the table. This is also an additional step to avoid wrongly merged blobs/cells
    # # Only for visualization purposes. Uncomment the lines below to visualize the table image showing blobs (text/digit locations), after deleting the original horizontal and vertical lines in the table to avoid over-merging of blobs incorrectly across columns and rows
    image_with_word_blobs = cv2.subtract(image_with_word_blobs, combined_image_dilated)
    # plt.imshow(image_with_word_blobs, cmap="gray")
    # plt.title("image_removing_previous_lines_areas_to_Avoid_merging_of_blobs_horizontally")
    # plt.show()


           
    ## STEP 5: ARRANGING THE DETECTED BLOBS, HERE CELLS, INTO ROWS AND COLUMNS READY FOR THE NEXT TRANSCRIPTION MODULE
    ## Here we draw contours around the remaining detected blobs in order to detect text in the table
    result = cv2.findContours(image_with_word_blobs, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    contours = result[0]
    # Original image of table in binarized format
    image_with_all_bounding_boxes = cv2.imread(save_path)

    # Make a copy of the original image to overlay contours without modifying the original
    table_img_bin_overlayed_with_contours = table_img_bin.copy()
    # Convert the grayscale image to RGB to support colored bounding boxes
    table_img_bin_overlayed_with_contours = cv2.cvtColor(table_img_bin_overlayed_with_contours, cv2.COLOR_GRAY2RGB)

    # Filter out smaller or larger bounding boxes from all the detected text contours. This is helpful to avoid overly large cells or small cells with no text. Remember to adjust these values for (min_cell_width_threshold, min_cell_height_threshold, max_cell_width_threshold, max_cell_height_threshold) based on the table structure in your specific case in the configutation.ini file.
    filtered_contours = filter_contours(contours, min_cell_width_threshold, min_cell_height_threshold, max_cell_width_threshold, max_cell_height_threshold)

    # Iterate over each contour in the new_contours list and draw bounding boxes
    for contour in filtered_contours:
        if contour is not None and len(contour) > 0:
            x, y, w, h = cv2.boundingRect(contour) # Dimensions of the rectangular countours surrounding the blobs. We term these as bounding boxes. 

            # # Adjust bounding box dimensions
            # Define increase factors for bounding box modification. We do this to give more room to the detected cell and avoid cutting the text in the cells.
            increase_factor_width = 0.07  # Increase width by 7%
            increase_factor_height = 0.20  # Increase height by 20%

            # Expand width while keeping it centered
            new_w = int(w * (1 + increase_factor_width))  # Increase width
            x = max(0, x - (new_w - w) // 2)  # Adjust x to keep center fixed

            # Expand height symmetrically
            new_h = int(h * (1 + increase_factor_height * 2))  # Increase height
            y = max(0, y - (new_h - h) // 2)  # Adjust y to keep center fixed

            # Ensure bounding box remains within valid image bounds
            x = max(0, x)
            y = max(0, y)
            w = max(1, new_w)  # Avoid zero or negative width
            h = max(1, new_h)  # Avoid zero or negative height
            
            # Draw the bounding box directly on the overlay image
            cv2.rectangle(table_img_bin_overlayed_with_contours, (x, y), (x + w, y + h), (0, 0, 255), 3)

    # # Only for visualization purposes. Display the image with bounding boxes using matplotlib
    # plt.imshow(table_img_bin_overlayed_with_contours)
    # plt.axis('off')  # Hide axis
    # plt.show()

    # Sort contours by y-coordinate
    contours_sorted = sorted(filtered_contours, key=lambda c: cv2.boundingRect(c)[1])



    ## STEP 6: REMOVE ANY NOISE (SUCH AS DOTS) FROM THE BINARIZED IMAGE OF THE TABLE TO PREPARE IT FOR THE NEXT TRANSCRIPTION MODULE
    # Apply the same dot removal to the **binarized image**
    table_img_bin[dots_to_remove == 255] = 255  # Set dots to white

    # Ensure the image is strictly binary
    table_img_bin[table_img_bin > 127] = 255  # Set everything above 127 to white
    table_img_bin[table_img_bin <= 127] = 0   # Set everything below 127 to black

    # Save the binary image for use later in detecting text
    save_dir = os.path.join(transient_transcription_output_dir, station)
    os.makedirs(save_dir, exist_ok=True)  # Ensure the directory exists
    save_path_without_dots = os.path.join(save_dir, 'table_binarized_without_dots.jpg')
    cv2.imwrite(save_path_without_dots, table_img_bin)

    # # Only for visualization purposes. show the binarized image of the table without any noise such as dots.
    # # Original image of table in binarized format without the dots
    # table_binarized_without_dots_file = cv2.imread(save_path_without_dots)
    # table_binarized_without_dots = table_binarized_without_dots_file.copy()
    # plt.imshow(table_binarized_without_dots)
    # plt.title("Filtered Image - Text Only - Without dots")
    # plt.show()

    table_for_overlaying_with_contours = table_img_bin.copy()
    # Convert the grayscale image to RGB to support colored bounding boxes
    table_for_overlaying_with_contours = cv2.cvtColor(table_for_overlaying_with_contours, cv2.COLOR_GRAY2RGB)
    detected_table_cells = [contours_sorted, image_with_all_bounding_boxes, table_for_overlaying_with_contours, table_original_image, full_detected_table_with_labels]


    return detected_table_cells

