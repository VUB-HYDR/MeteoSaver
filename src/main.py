import os
from datetime import datetime

# Import all the required modules 
from src.image_preprocessing_module import *
from src.table_detection_model import *
from src.transcription_model import *
from src.post_processing_module import *
from src.selection_and_conversion import *
from src.validation import *

## ***NEW
# from src.template_matching import *
# from src.crop_border import *

# Setting up the current working directory; for both the input and output folders
cwd = os.getcwd()

# Home directory for original images for only one station for code testing purposes
full_datadir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_5_RawData\21_5_1_Precipitation'

station = 'DUMMY_FOLDER' # Here, I am using a dummy folder. However, this will be updated to a loop for all the data sheets. Within this dummy folder, I place some climate data sheets.
datadir = os.path.join(full_datadir, station)

# Home directory for preprocessed data from transcription process
preprocessed_data_dir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_0_Preprocessing_data'
preprocessed_data_dir_station = os.path.join(preprocessed_data_dir, station) # Folder for every station (for preprocessed data)
os.makedirs(preprocessed_data_dir_station, exist_ok=True) # Create the directory if it doesn't exist

# Home directory for postprocessed data from tracnscription process
postprocessed_data_dir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_1_Postprocessing_data'
postprocessed_data_dir_station = os.path.join(postprocessed_data_dir, station) # Folder for every station (for postprocessed data)
os.makedirs(postprocessed_data_dir_station, exist_ok=True) # Create the directory if it doesn't exist

# Station data - original photographs/scans of archived datasheets
station_data = glob.glob(datadir+ '/'+ str(station) +'_*') # a [] number of images from one station.
filenames = [os.path.basename(file) for file in station_data]

for month in range(len(station_data)):
    
    # Monthly Data for specific station
    month_data = station_data[month]
    month_filename = filenames[month] #restore naming of output files with station metadata

    # Module 1: Pre-processing the original images
    preprocessed_image = image_preprocessing(month_data)[0]
    original_image = image_preprocessing(month_data)[1]
    # Table detection
    detected_table_cells = table_detection(preprocessed_image, original_image, clip_up = 430, clip_down = 270, clip_left = 200, clip_right = 150) # Here the clip_up, clip_down, clip_left, and clip_right ensure clipping of the HEADERS and ROW LABELS (Date & Pentad no. in our case) from the entire detected table (table detected using ML). Adjust this to your case study. Incase you would like to maintain the full table, set clip_up, clip_down, clip_left, clip_right = 0

    # Module 2: Transcription / Handwritten Text Recognition
    start_time = datetime.now() # Start recording transcribing time
    ocr_model = 'Tesseract-OCR' # Selected OCR model out of: Tesseract-OCR, EasyOCR, PaddleOCR
    transcribed_table = transcription(detected_table_cells, ocr_model, no_of_rows = 43, no_of_columns = 24, min_cell_width_threshold = 50, max_cell_width_threshold = 200, min_cell_height_threshold = 28, max_cell_height_threshold = 90)  # Adjust these values based on the table structure in your specific case. The min_cell_width_threshold, max_cell_width_threshold, min_cell_height_threshold, and max_cell_height_threshold are used to filter out smaller or larger bounding boxes from all the detected text contours. this is helpful to avoid overly large cells or small cells with no text
    merge_excel_files(f'src\output\Top_Excel_with_OCR_Results.xlsx', f'src\output\Midpoint_Excel_with_OCR_Results.xlsx', f'{preprocessed_data_dir_station}\{month_filename}_preprocessed.xlsx', 1, 46) # this prioritizes the top coordinates of the bounding box to the mid point coordinates when placing the transcribed data into an excel sheet. But considers the best placement for both as double check. Here the 1 represent the first row and the 46 represents the last possible row to perform the merging of the excel files. In our case we have 43 rows with data + 3 header rows = 46.
    end_time=datetime.now() # print total runtime of the code
    print('Duration of transcribing: {}'.format(end_time - start_time))

    # Module 3: Quality Assessment and Quality Control (QA/QC). This is very specific to table structure and thus should be adapted according to your specific needs. Nevertheless, it provides a good example of how to improve the accuracy of transcribed data by performing relavant QA/QC checks
    start_time = datetime.now() # Start recording post-processing time
    post_processed_data = post_processing(f'{preprocessed_data_dir_station}\{month_filename}_preprocessed.xlsx', postprocessed_data_dir_station, month_filename)
    # post_processed_data = post_processing(f'src\output\Top_Excel_with_OCR_Results.xlsx', postprocessed_data_dir_station, month_filename)
    end_time=datetime.now() # print total runtime of the code for post-processing which includes the Quality Assessment and Quality Control (QA/QC) checks
    print('Duration of post-processing: {}'.format(end_time - start_time))
    


# Module 4: Selection of confirmed data (after Quality Control) and conversion to Station Exchange Format (SEF)
selected_data_dir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_3_Final_refined_daily_station_data'
selected_data_dir_station = os.path.join(selected_data_dir, station) # Folder for every station (for postprocessed data)
os.makedirs(selected_data_dir_station, exist_ok=True) # Create the directory if it doesn't exist
selected_and_converted_data = select_and_convert_postprocessed_data(postprocessed_data_dir_station, selected_data_dir_station, station)

# Validation: This is step is only done for already manually transcibed data
manually_transcribed_data_dir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_1_Postprocessing_data\Manually_transcribed_data'
validation_dir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_2_Validation'
validate(manually_transcribed_data_dir, postprocessed_data_dir_station, validation_dir, station)

