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


# Directory for the original images (data sheets) to be transcribed
# ***Change this directory. Below we are testing this transcribing model on a folder with 10 sample images, here called 'data'
# images_folder = os.path.join(cwd, 'data') #folder containing images and and template guides file
# sample_images = os.path.join(images_folder, '10_sample_different_images') # 10 sample images
# Home directory for original images for only one station for code testing purposes
datadir = r'C:\Users\dmuheki\OneDrive - Vrije Universiteit Brussel\PhD_Derrick_Muheki\21_Research\21_5_RawData\21_5_1_Precipitation\203'
station_203 = glob.glob(datadir+'/IMG_*') # a [] number of images from one station. Tasked the code on 203[124]
# filenames=[]
# for file in station_203:
#     filename=os.path.basename(file)
#     filenames.append(filename)


# TRIAL ON ONE TEST IMAGE FROM THE FOLDER
## This will be replaced with a 'for' loop after testing all the functions
# one_test_image =  os.path.join(sample_images, '205_198701_SF_YAN.JPG')
# # Module 1: Pre-processing the original images
# preprocessed_image = image_preprocessing(one_test_image)[0]
# original_image = image_preprocessing(one_test_image)[1]

# # Module 2: Table detection
# detected_table_cells = table_detection(preprocessed_image, original_image)

# # Module 3: Transcription / Handwritten Text Recognition
# start_time = datetime.now() # Start recording transcribing time
# ocr_model = 'Tesseract-OCR' # Selected OCR model out of: Tesseract-OCR, EasyOCR, PaddleOCR
# transcribed_table = transcription(detected_table_cells, ocr_model)
# end_time=datetime.now() # print total runtime of the code
# print('Duration of transcribing: {}'.format(end_time - start_time))

# TRIAL LOOP ON IMAGES FROM ONE STATION 
for month_data in [station_203]:

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
    post_processed_data = post_processing('Excel_with_OCR_Results.xlsx')
    end_time=datetime.now() # print total runtime of the code
    print('Duration of post-processing: {}'.format(end_time - start_time))







    # # Module 4: Post-processing
    # start_time = datetime.now() # Start recording post-processing time
    # post_processed_data = post_processing('Excel_with_OCR_Results.xlsx')
    # end_time=datetime.now() # print total runtime of the code
    # print('Duration of post-processing: {}'.format(end_time - start_time))


# new_version_of_file = 'quality_controlled_data_table_copy.xlsx' # To save a new copy after quality control

# workbook = openpyxl.load_workbook('Excel_with_OCR_Results.xlsx')
# worksheet = workbook.active # To select the first worksheet of the workbook without requiring its name

# # Create a copy of the oriinal file
# shutil.copy2(workbook, new_version_of_file)
# new_workbook = openpyxl.load_workbook(new_version_of_file)
# new_worksheet = new_workbook.active # To select the first worksheet of the workbook without requiring its name

# # Module 4: Post-processing
# start_time = datetime.now() # Start recording post-processing time
# post_processed_data = post_processing('Excel_with_OCR_Results.xlsx')
# end_time=datetime.now() # print total runtime of the code
# print('Duration of post-processing: {}'.format(end_time - start_time))

import pandas as pd

# ## WORK'S REALLY WELL
def merge_excel_files(file1, file2, output_file, start_row, end_row):
    # Load the Excel files into DataFrames, ensuring they include headers if present
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Load headers separately if you need to prepend them later
    headers = pd.read_excel(file1, nrows=3)  # Read only the first three rows for headers

    # # Print DataFrame indices to understand their structure
    # print("DF1 Index:", df1.index)
    # print("DF2 Index:", df2.index)

    # # Check the first few rows to understand how the data looks
    # print("DF1 head:", df1.head())
    # print("DF2 head:", df2.head())

    # If the indices are not simple integers or do not align with Excel rows as expected,
    # you might need to reset them or adjust how the Excel file is being read (e.g., `index_col=None`)
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)

    # Convert start_row and end_row to zero-based index for Python
    start_idx = start_row - 1  # Convert 1-based index to 0-based
    end_idx = end_row -1    # Convert 1-based index to 0-based

    # Slice to only include the range from start_idx to end_idx
    df1 = df1.iloc[start_idx:end_idx+1]
    df2 = df2.iloc[start_idx:end_idx+1]

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

    # Write the merged DataFrame to a new Excel file
    # merged_df.to_excel(output_file)

    # Prepend headers if needed
    final_df = pd.concat([headers, merged_df], ignore_index=True)

    # Write the merged DataFrame to a new Excel file without the index
    final_df.to_excel(output_file, index=False, header=None)  # Set header=None if headers are manually handled

merge_excel_files('Midpoint_Excel_with_OCR_Results.xlsx', 'Top_Excel_with_OCR_Results.xlsx', 'Cross_checked_Excel_with_OCR_Results.xlsx', 4, 48) # this prioritizes the mid point coordinates of the bounding box to the top coordinates when placing the transcribed data into an excel sheet. But considers the best placement for both as double check.
#merge_excel_files('Top_Excel_with_OCR_Results.xlsx', 'Midpoint_Excel_with_OCR_Results.xlsx', 'Cross_checked_Excel_with_OCR_Results.xlsx', 4, 48)
# merge_excel_files('Top_Excel_with_OCR_Results.xlsx', 'Bottom_Excel_with_OCR_Results.xlsx', 'Cross_checked_Excel_with_OCR_Results.xlsx', 4, 48)


# def merge_excel_files(file1, file2, output_file, start_row, end_row):
#     # Load the Excel files into DataFrames, ensuring they include headers if present
#     df1 = pd.read_excel(file1)
#     df2 = pd.read_excel(file2)

#     # Load headers separately if you need to prepend them later
#     headers = pd.read_excel(file2, nrows=3)  # Read only the first three rows for headers

#     # # Print DataFrame indices to understand their structure
#     # print("DF1 Index:", df1.index)
#     # print("DF2 Index:", df2.index)

#     # # Check the first few rows to understand how the data looks
#     # print("DF1 head:", df1.head())
#     # print("DF2 head:", df2.head())

#     # If the indices are not simple integers or do not align with Excel rows as expected,
#     # you might need to reset them or adjust how the Excel file is being read (e.g., `index_col=None`)
#     df1 = df1.reset_index(drop=True)
#     df2 = df2.reset_index(drop=True)

#     # Convert start_row and end_row to zero-based index for Python
#     start_idx = start_row - 1  # Convert 1-based index to 0-based
#     end_idx = end_row - 1      # Convert 1-based index to 0-based

#     # Slice to only include the range from start_idx to end_idx
#     df1 = df1.iloc[start_idx:end_idx+1]
#     df2 = df2.iloc[start_idx:end_idx+1]

#     # Initialize a new DataFrame to hold merged results
#     merged_df = pd.DataFrame(index=df1.index, columns=df1.columns)

#     # Iterate over rows by index (assuming the indices are aligned)
#     for idx in df1.index:
#         for col in df1.columns:
#             val1 = df1.at[idx, col]
#             val2 = df2.at[idx, col]
#             # Simple merge logic: prefer non-empty values from df1, then df2
#             if pd.notna(val1):
#                 merged_df.at[idx, col] = val1
#             else:
#                 merged_df.at[idx, col] = val2

#     # Write the merged DataFrame to a new Excel file
#     # merged_df.to_excel(output_file)
#     # Ensure headers and data columns are aligned
#     if headers.shape[1] != merged_df.shape[1]:
#         # Add missing columns as NaN or with a placeholder to headers or merged_df
#         max_cols = max(headers.shape[1], merged_df.shape[1])
#         headers = headers.reindex(columns=range(max_cols), fill_value='')
#         merged_df = merged_df.reindex(columns=range(max_cols), fill_value='')

#     # Prepend headers if needed
#     final_df = pd.concat([headers, merged_df], ignore_index=True)

#     # Write the merged DataFrame to a new Excel file without the index
#     final_df.to_excel(output_file, index=False, header=False)  # Set header=None if headers are manually handled



# def merge_excel_files(file1, file2, output_file, start_row, end_row):
#     # Load the Excel files into DataFrames, ensuring they include headers if present
#     df1 = pd.read_excel(file1)
#     df2 = pd.read_excel(file2)

#     # Load headers separately if you need to prepend them later
#     headers = pd.read_excel(file1,header = None, nrows=3)  # Read only the first three rows for headers

#     # # Print DataFrame indices to understand their structure
#     # print("DF1 Index:", df1.index)
#     # print("DF2 Index:", df2.index)

#     # # Check the first few rows to understand how the data looks
#     # print("DF1 head:", df1.head())
#     # print("DF2 head:", df2.head())

#     # If the indices are not simple integers or do not align with Excel rows as expected,
#     # you might need to reset them or adjust how the Excel file is being read (e.g., `index_col=None`)
#     df1 = df1.reset_index(drop=True)
#     df2 = df2.reset_index(drop=True)

#     # Convert start_row and end_row to zero-based index for Python
#     start_idx = start_row - 1  # Convert 1-based index to 0-based
#     end_idx = end_row - 1      # Convert 1-based index to 0-based

#     # Slice to only include the range from start_idx to end_idx
#     df1 = df1.iloc[start_idx:end_idx+1]
#     df2 = df2.iloc[start_idx:end_idx+1]

#     # Initialize a new DataFrame to hold merged results
#     merged_df = pd.DataFrame(index=df1.index, columns=df1.columns)

#     # Iterate over rows by index (assuming the indices are aligned)
#     for idx in df1.index:
#         for col in df1.columns:
#             val1 = df1.at[idx, col]
#             val2 = df2.at[idx, col]
#             # Simple merge logic: prefer non-empty values from df1, then df2
#             if pd.notna(val1):
#                 merged_df.at[idx, col] = val1
#             else:
#                 merged_df.at[idx, col] = val2

#     # Write the merged DataFrame to a new Excel file
#     merged_df.to_excel(output_file)

#     # Prepend headers if needed
#     final_df = pd.concat([headers, merged_df], ignore_index=True)

#     # Write the merged DataFrame to a new Excel file without the index
#     final_df.to_excel(output_file, index=False, header=False)  # Set header=None if headers are manually handled




# def merge_excel_files(file1, file2, output_file, start_row, end_row):
#     # Load the Excel files into DataFrames, assuming headers are in the first three rows
#     df1 = pd.read_excel(file1, header=None, skiprows=3)  # Skip first 3 rows if they are headers
#     df2 = pd.read_excel(file2, header=None, skiprows=3)  # Adjust as necessary

#     # Load headers separately if you need to prepend them later
#     # headers = pd.read_excel(file1, nrows=3)  # Read only the first three rows for headers

#     # Print DataFrame indices to check structure (optional)
#     print("DF1 Index:", df1.index)
#     print("DF2 Index:", df2.index)

#     # Resetting index to ensure it's a simple 0-based integer index
#     df1 = df1.reset_index(drop=True)
#     df2 = df2.reset_index(drop=True)

#     # Adjust index for 0-based in Python
#     start_idx = start_row - 4  # Adjust start row to account for skipped rows
#     end_idx = end_row - 4      # Adjust end row similarly

#     # Slicing DataFrames
#     df1 = df1.iloc[start_idx:end_idx+1]
#     df2 = df2.iloc[start_idx:end_idx+1]

#     # Initialize a new DataFrame to hold merged results
#     merged_df = pd.DataFrame(index=df1.index, columns=df1.columns)

#     # Merge logic
#     for idx in df1.index:
#         for col in df1.columns:
#             val1 = df1.at[idx, col]
#             val2 = df2.at[idx, col]
#             if pd.notna(val1):
#                 merged_df.at[idx, col] = val1
#             else:
#                 merged_df.at[idx, col] = val2

#     # Write the merged DataFrame to a new Excel file
#     # merged_df.to_excel(output_file)
    
#     # Prepend headers if needed
#     #final_df = pd.concat([headers, merged_df], ignore_index=True)

#     # Write the merged DataFrame to a new Excel file without the index
#     merged_df.to_excel(output_file, index=False, header=None)  # Set header=None if headers are manually handled


# def merge_excel_files(file1, file2, output_file, start_row, end_row):
#     # Load the Excel files into DataFrames
#     df1 = pd.read_excel(file1)
#     df2 = pd.read_excel(file2)

#     # Load headers separately
#     headers = pd.read_excel(file2, nrows=3)  # Assuming headers are the first 3 rows

#     # Reset index for proper alignment
#     df1 = df1.reset_index(drop=True)
#     df2 = df2.reset_index(drop=True)

#     # Adjust indices for 0-based Python indexing
#     start_idx = start_row - 1
#     end_idx = end_row - 1

#     # Slice the DataFrames
#     df1 = df1.iloc[start_idx:end_idx+1]
#     df2 = df2.iloc[start_idx:end_idx+1]

#     # Create a merged DataFrame
#     merged_df = pd.DataFrame(index=df1.index, columns=df1.columns)
#     for idx in df1.index:
#         for col in df1.columns:
#             val1 = df1.at[idx, col]
#             val2 = df2.at[idx, col]
#             merged_df.at[idx, col] = val1 if pd.notna(val1) else val2

#     # Align columns between headers and merged data
#     if headers.shape[1] != merged_df.shape[1]:
#         # Extend the DataFrame with missing columns filled with empty strings
#         new_cols = list(set(headers.columns).union(set(merged_df.columns)))
#         headers = headers.reindex(columns=new_cols, fill_value='')
#         merged_df = merged_df.reindex(columns=new_cols, fill_value='')

#     # Concatenate headers and data
#     final_df = pd.concat([headers, merged_df], ignore_index=True)

#     # Write to Excel without index
#     final_df.to_excel(output_file, index=False, header=False)


# Usage
merge_excel_files('quality_controlled_data_table_copy_better_but_still_a_little_messed_up_rows.xlsx', 'quality_controlled_data_table_copy.xlsx', 'cross_checked_quality_controlled_data_table_copy.xlsx', 3, 48)

# Module 4: Post-processing
start_time = datetime.now() # Start recording post-processing time
post_processed_data = post_processing('Cross_checked_Excel_with_OCR_Results.xlsx')
end_time=datetime.now() # print total runtime of the code
print('Duration of post-processing: {}'.format(end_time - start_time))