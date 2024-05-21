import os
from datetime import datetime

# Import all the required modules 
from src.image_preprocessing_module import *
from src.table_detection_model import *
from src.transcription_model import *
from src.post_processing_module import *

## ***NEW
from src.template_matching import *
from src.crop_border import *

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
# station_data = glob.glob(datadir+'/IMG_*') # a [] number of images from one station. Tasked the code on 203[124]
station_data = glob.glob(datadir+ '/'+ str(station) +'_*') # a [] number of images from one station. Tasked the code on 203[124]
filenames = [os.path.basename(file) for file in station_data]

for month in range(len(station_data)):
    
    # Monthly Data for specific station
    month_data = station_data[month]

    # Module 1: Pre-processing the original images
    preprocessed_image = image_preprocessing(month_data)[0]
    original_image = image_preprocessing(month_data)[1]

    # Module 2: Table detection
    detected_table_cells = table_detection(preprocessed_image, original_image)

    # Module 3: Transcription / Handwritten Text Recognition
    start_time = datetime.now() # Start recording transcribing time
    ocr_model = 'Tesseract-OCR' # Selected OCR model out of: Tesseract-OCR, EasyOCR, PaddleOCR
    transcribed_table = transcription(detected_table_cells, ocr_model)
    end_time=datetime.now() # print total runtime of the code
    print('Duration of transcribing: {}'.format(end_time - start_time))

    # Module 4: Post-processing
    start_time = datetime.now() # Start recording post-processing time
    month_filename = filenames[month] #restore naming of output files with station metadata
    merge_excel_files(f'src\output\Midpoint_Excel_with_OCR_Results.xlsx', f'src\output\Top_Excel_with_OCR_Results.xlsx', f'{preprocessed_data_dir_station}\{month_filename}_preprocessed.xlsx', 4, 48) # this prioritizes the mid point coordinates of the bounding box to the top coordinates when placing the transcribed data into an excel sheet. But considers the best placement for both as double check.
    post_processed_data = post_processing(f'{preprocessed_data_dir_station}\{month_filename}_preprocessed.xlsx', postprocessed_data_dir_station, month_filename)
    end_time=datetime.now() # print total runtime of the code
    print('Duration of post-processing: {}'.format(end_time - start_time))


    # ## Validation. This will be implemetend as the last step for selected data
    # # Accuracy check# Example usage
    # new_file1 = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_1_Postprocessing_data\DUMMY_FOLDER\DUMMY_FOLDER_196905_SF.JPG_post_processed.xlsx'  # Workbook with highlighted cells
    # file2 = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_1_Postprocessing_data\DUMMY_FOLDER\IMG_1361.JPG_manually_entered_temperatures.xlsx'  # Workbook with correct values
    # compare_workbooks(new_file1, file2)

    # old_file1 = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_1_Postprocessing_data\DUMMY_FOLDER\IMG_1361.JPG_post_processed.xlsx'  # Workbook with highlighted cells
    # file2 = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_6_Analysis\21_6_1_Postprocessing_data\DUMMY_FOLDER\IMG_1361.JPG_manually_entered_temperatures.xlsx'  # Workbook with correct values
    # compare_workbooks(old_file1, file2)



