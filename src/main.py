import os
import configparser
import multiprocessing as mp
from datetime import datetime
import glob

# Import all the required modules
from image_preprocessing_module import *
from table_and_cell_detection_model import *
from transcription import *
from quality_assessment_and_quality_control import *
from data_formatting_and_upload import *
from validation import *

# Module 1: Configuration
# Load settings from user configurations. See configuration.ini file in this repository
config = configparser.ConfigParser()
# Adjust path to config.ini (since it's placed in the root directory)
config_file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'configuration.ini')
config.read(config_file_path)

# Get run_mode and number of processors from configuration file
run_mode = config['General']['run_mode']
num_processors = int(config['General'].get('num_processors', 1))

# Set up directories from the configuration file
full_datadir = config['Directories']['full_datadir'] # Directory to all the images/scans of the hydroclimatic data sheets
pre_QA_QC_transcribed_hydroclimate_data_dir = config['Directories']['pre_QA_QC_transcribed_hydroclimate_data_dir'] # Directory where pre-QA/QC transcribed data is stored
post_QA_QC_transcribed_hydroclimate_data_dir = config['Directories']['post_QA_QC_transcribed_hydroclimate_data_dir'] # Directory where post-QA/QC transcribed data is stored
final_refined_daily_hydroclimate_data_dir = config['Directories']['final_refined_daily_hydroclimate_data_dir'] # Directory for the final refined daily hydroclimate data (after all quality checks)
manually_transcribed_data_dir = config['Directories']['manually_transcribed_data_dir'] # Directory for manually transcribed data (used for validation)
validation_dir = config['Directories']['validation_dir'] # Directory for validation results comparing manually transcribed and the MeteoSaver transcribed data 
transient_transcription_output_dir = config['Directories']['transient_transcription_output_dir'] # Directory to store transient transcription output during processing
metadata_file_path = config['Directories']['metadata_file_path'] # Directory for all the stations metadata

# Get all folder names (Station Numbers) within full_datadir
all_stations = [folder for folder in os.listdir(full_datadir) if os.path.isdir(os.path.join(full_datadir, folder))]

# Function to process each station's data
def process_station(station):
    datadir = os.path.join(full_datadir, station)
    pre_QA_QC_transcribed_hydroclimate_data_dir_station = os.path.join(pre_QA_QC_transcribed_hydroclimate_data_dir, station)
    post_QA_QC_transcribed_hydroclimate_data_dir_station = os.path.join(post_QA_QC_transcribed_hydroclimate_data_dir, station)
    
    # Ensure directories exist
    os.makedirs(pre_QA_QC_transcribed_hydroclimate_data_dir_station, exist_ok=True)
    os.makedirs(post_QA_QC_transcribed_hydroclimate_data_dir_station, exist_ok=True)
    
    # Search for all files in the `datadir` that start with the station number followed by an underscore ('_') and any other characters, and return a list of all matching file paths
    station_data = glob.glob(os.path.join(datadir, f"{station}_*")) 
    # OPTIONAL (Comment these two lines below if unnecessary): Filter the station_data and filenames to include only files with 'SF' in their names. Here the SF for standard format for sheets. This to filter our sheets with HD in their name, which stands for hand-drawn format, in their name as these were not standard formatted sheets and manually drawn by the observer.
    filenames = [os.path.basename(file) for file in station_data if 'SF' in os.path.basename(file)]

    for month in range(len(filenames)):
        month_data = station_data[month]
        month_filename = filenames[month]

        # Perform Pre-processing, Transcription, QA/QC, and Post-processing
        try:
            # Module 2: Image pre-processing
            image_in_grayscale, binarized_image, original_image = image_preprocessing(month_data)

            # Module 3: Table and cell detection 
            detected_table_and_cells = table_and_cell_detection(image_in_grayscale, binarized_image, original_image, station, month_filename, transient_transcription_output_dir,
                                                   clip_up = int(config['TableAndCellDetection']['clip_up']), 
                                                   clip_down = int(config['TableAndCellDetection']['clip_down']),
                                                   clip_left = int(config['TableAndCellDetection']['clip_left']),
                                                   clip_right = int(config['TableAndCellDetection']['clip_right']),
                                                   max_table_width = int(config['TableAndCellDetection']['max_table_width']),
                                                   max_table_height = int(config['TableAndCellDetection']['max_table_height']),
                                                   min_cell_width_threshold=int(config['TableAndCellDetection']['min_cell_width_threshold']),
                                                   max_cell_width_threshold=int(config['TableAndCellDetection']['max_cell_width_threshold']),
                                                   min_cell_height_threshold=int(config['TableAndCellDetection']['min_cell_height_threshold']),
                                                   max_cell_height_threshold=int(config['TableAndCellDetection']['max_cell_height_threshold']),
                                                   space_height_threshold=int(config['TableAndCellDetection']['space_height_threshold']), 
                                                   space_width_threshold=int(config['TableAndCellDetection']['space_width_threshold']), 
                                                   max_cell_height_per_box=int(config['TableAndCellDetection']['max_cell_height_per_box']), 
                                                   no_of_rows=int(config['TableAndCellDetection']['no_of_rows']), 
                                                   no_of_columns=int(config['TableAndCellDetection']['no_of_columns']))
            
            # Module 4: Transcription
            start_time = datetime.now()
            ocr_model = config['Transcription']['ocr_model'] # Selected OCR/HTR model
            # Incase of Tesseract
            # Ensure that the tesseract path is set correctly for your local system
            tesseract_path = config['Transcription']['tesseract_path']
            # Set TESSDATA_PREFIX to the system's tessdata directory (for system-wide language files)
            system_tessdata_dir = config['Transcription']['system_tessdata_dir']
            os.environ["TESSDATA_PREFIX"] = system_tessdata_dir
            transcribed_table = transcription(detected_table_and_cells, ocr_model, tesseract_path, transient_transcription_output_dir, pre_QA_QC_transcribed_hydroclimate_data_dir_station, station, month_filename,
                                              no_of_rows=int(config['TableAndCellDetection']['no_of_rows']),
                                              no_of_columns=int(config['TableAndCellDetection']['no_of_columns']),
                                              no_of_rows_including_headers=int(config['TableAndCellDetection']['no_of_rows_including_headers']))
            
            end_time = datetime.now()

            print(f'Duration of transcribing: {end_time - start_time}')
            
            # Module 5: Quality assessment and Quality Control
            qa_qc_checked_data = qa_qc(transcribed_table, station, transient_transcription_output_dir, post_QA_QC_transcribed_hydroclimate_data_dir_station, month_filename,
                                        max_temperature_threshold = float(config['QAQC']['max_temperature_threshold']),
                                        min_temperature_threshold = float(config['QAQC']['min_temperature_threshold']),
                                        decimal_places = int(config['QAQC']['decimal_places']),
                                        uncertainty_margin = float(config['QAQC']['uncertainty_margin']),
                                        header_rows = int(config['QAQC']['header_rows']),
                                        multi_day_totals = config.getboolean('QAQC', 'multi_day_totals'),
                                        multi_day_averages = config.getboolean('QAQC', 'multi_day_averages'),
                                        max_days_for_multi_day_total = int(config['QAQC']['max_days_for_multi_day_total']),
                                        multi_day_totals_rows = list(map(int, config['QAQC']['multi_day_totals_rows'].split(','))),
                                        final_totals_rows = list(map(int, config['QAQC']['final_totals_rows'].split(','))),
                                        excluded_rows = list(map(int, config['QAQC']['excluded_rows'].split(','))),
                                        excluded_columns = list(map(int, config['QAQC']['excluded_columns'].split(','))),
                                        columns_to_check = config['QAQC']['columns_to_check'].split(','),
                                        columns_to_check_with_extra_variable = config['QAQC']['columns_to_check_with_extra_variable'].split(','))
        
        except Exception as e:
            print(f"Error processing {month_filename}: {e}")
            continue

    # Module 6: Data formatting and Upload
    final_refined_daily_hydroclimate_data_dir_station = os.path.join(final_refined_daily_hydroclimate_data_dir, station)
    os.makedirs(final_refined_daily_hydroclimate_data_dir_station, exist_ok=True)
    
    data_formatting(post_QA_QC_transcribed_hydroclimate_data_dir_station, final_refined_daily_hydroclimate_data_dir_station, metadata_file_path, station, 
                    date_column = config['DataFormatting']['date_column'].strip(),
                    columns_to_check = config['QAQC']['columns_to_check'].split(','),
                    header_rows = int(config['QAQC']['header_rows']),
                    multi_day_totals = config.getboolean('QAQC', 'multi_day_totals'),
                    multi_day_averages = config.getboolean('QAQC', 'multi_day_averages'),
                    excluded_rows = list(map(int, config['QAQC']['excluded_rows'].split(','))),
                    additional_excluded_rows = list(map(int, config['QAQC']['additional_excluded_rows'].split(','))),
                    final_totals_rows = list(map(int, config['QAQC']['final_totals_rows'].split(','))),
                    uncertainty_margin = float(config['QAQC']['uncertainty_margin']))

    # Extra module: Validation
    validation_dir_station = os.path.join(validation_dir, station)
    manually_transcribed_data_dir_station = os.path.join(manually_transcribed_data_dir, station)
    # Ensure directories exist
    os.makedirs(validation_dir_station, exist_ok=True)
    os.makedirs(manually_transcribed_data_dir_station, exist_ok=True)

    validate(manually_transcribed_data_dir_station, post_QA_QC_transcribed_hydroclimate_data_dir_station, validation_dir_station, station,
             date_column = config['DataFormatting']['date_column'].strip(),
             columns_to_check = config['QAQC']['columns_to_check'].split(','),
             header_rows = int(config['QAQC']['header_rows']),
             multi_day_totals = config.getboolean('QAQC', 'multi_day_totals'),
             multi_day_averages = config.getboolean('QAQC', 'multi_day_averages'),
             excluded_rows = list(map(int, config['QAQC']['excluded_rows'].split(','))),
             additional_excluded_rows = list(map(int, config['QAQC']['additional_excluded_rows'].split(','))),
             final_totals_rows = list(map(int, config['QAQC']['final_totals_rows'].split(','))),
             uncertainty_margin = float(config['QAQC']['uncertainty_margin']))

if __name__ == '__main__':
    if run_mode == 'hpc':
        # HPC Mode: Parallel Processing
        print(f"Running in HPC mode with {num_processors} processors.")
        with mp.Pool(num_processors) as pool:
            pool.map(process_station, all_stations)
    else:
        # Local Mode: Sequential Processing
        print("Running in Local mode.")
        for station in all_stations:
            process_station(station)
