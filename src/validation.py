import os
import openpyxl
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
import matplotlib.dates as mdates

def load_data(filepath):
    """Load data from the given Excel file."""
    return pd.read_excel(filepath)


# Function to extract year, month from filename
def extract_date_from_filename(filename):
    '''
    # Extract the date from the filename

    Parameters
    --------------
    filename: filename of the postprocessed data. These are in the structure/format "STN_YYYYMM_SF" or "STN_YYYYMM_HD" using the data inventory, where:

            STN is the three digit station number,
            YYYY is the year
            MM is the month,
            SF represents Standard Format
            HD represents a hand drawn form / or photocopied form of the standard format

    Returns
    -------------- 
    year, month : Year (YYYY), Month (MM)
    
    '''

    match = re.search(r'_(\d{4})(\d{2})_', filename)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        return year, month
    return None, None


def is_highlighted_green(cell, color):
    ''' Checks if a cell is highlighted 'GREEN' during the earlier post processing steps that symbolizes that this data was confirmed 

    Parameters
    --------------   
    cell: coordinates of cell to check
    color: highlighted color

    Returns
    -------------- 
    boolean: 1 (True) if the cell is highlighted with GREEN (i.e. confirmed in Quality Control)

    '''

    fill = cell.fill.start_color
    if isinstance(fill, openpyxl.styles.colors.Color):
        return fill.rgb == color
    return False



def select_and_convert_postprocessed_data(filepath, date_column, columns_to_check, header_rows, multi_day_totals, multi_day_averages, excluded_rows, additional_excluded_rows, final_totals_rows):

    ''' Selects the confirmed data (from the quality control) and converts the excel file to a format ready to be converted to the Station Exchange Format
    Parameters
    --------------   
    filepath: path of QA/QC checked files

    Returns
    -------------- 
    Excel files with selected (confirmed) transcribed data -> in new format 
    Timeseries plot of max, min and avg temperature (confirmed) at particular station
    
    '''

    # Extract year and month from the filename
    filename = os.path.basename(filepath)  # Get the filename from the full file path
    year, month = extract_date_from_filename(filename)  # Assuming `extract_date_from_filename` is a utility that parses the year/month from the filename

    if year and month:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active

    if not (year and month):
        raise ValueError(f"Could not extract year and month from filename: {filename}")


    # List to hold all data
    data = [] # All the variables

    # Lists to hold all data for each temperature type
    data_max = []
    data_min = []
    data_avg = []

    # Rows to exclude. Adjust these according to your specific sheet
    if multi_day_totals and not multi_day_averages:
        excluded_rows =  excluded_rows # These include titles and multi-day (e.g. 5/6) day totals
    if multi_day_totals and multi_day_averages:
        excluded_rows = excluded_rows + additional_excluded_rows # including both multi day totals and averages
    if not multi_day_totals:
        excluded_rows = final_totals_rows # Exlude only the final totals

    # Convert the day column letter and temperature columns to numeric indices
    date_column_idx = ord(date_column) - ord('A') + 1  # Convert 'B' -> 2  # Date
    column_indices = [ord(col.strip()) - ord('A') + 1 for col in columns_to_check] # Max, min and average temperatures 

    # Now `column_indices` will contain [4, 5, 6] for 'D', 'E', 'F'
    max_temp_column_idx = column_indices[0]  # 'D' column index -> Maximum temperature
    min_temp_column_idx = column_indices[1]  # 'E' column index -> Minimum temperature
    avg_temp_column_idx = column_indices[2]  # 'F' column index -> Avergae temperature

    total_transcribed_cells = 0 


    # Extract data from rows and columns, excluding specific rows.  
    for row_num in range(header_rows+1, worksheet.max_row+1): #Here this represents Max, Min and Average Temperatures
        if row_num not in excluded_rows: 
            day_cell = worksheet.cell(row=row_num, column=date_column_idx)  # Assuming the day is in the first column
            max_temperature_cell = worksheet.cell(row=row_num, column=max_temp_column_idx)  # Column for Max Temperature
            min_temperature_cell = worksheet.cell(row=row_num, column=min_temp_column_idx)  # Column for Min Temperature
            average_temperature_cell = worksheet.cell(row=row_num, column=avg_temp_column_idx)  # Column for Avg Temperature

            # Count the total transcribed values, including even the non-confirmed (GREEN) ones
            total_transcribed_cells += sum(cell.value is not None for cell in [max_temperature_cell, min_temperature_cell, average_temperature_cell])

            if day_cell.value is not None and day_cell.value.isdigit():
                day = int(day_cell.value)
                max_temperature = max_temperature_cell.value if is_highlighted_green(max_temperature_cell, 'FF6DCD57') else 'NaN'
                min_temperature = min_temperature_cell.value if is_highlighted_green(min_temperature_cell, 'FF6DCD57') else 'NaN'
                average_temperature = average_temperature_cell.value if is_highlighted_green(average_temperature_cell, 'FF6DCD57') else 'NaN'
                
                data.append([year, month, day, max_temperature, min_temperature, average_temperature])

                data_max.append([year, month, day, max_temperature])
                data_min.append([year, month, day, min_temperature])
                data_avg.append([year, month, day, average_temperature])


    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=["Year", "Month", "Day", "Max_Temperature", "Min_Temperature", "Avg_Temperature"])


    # Generate a complete date range for each year and month combination
    years_months = df[['Year', 'Month']].drop_duplicates()
    complete_data = []

    for _, row in years_months.iterrows():
        year = row['Year']
        month = row['Month']
        num_days = pd.Period(f'{year}-{month}').days_in_month
        for day in range(1, num_days + 1):
            complete_data.append([year, month, day])

    complete_df = pd.DataFrame(complete_data, columns=["Year", "Month", "Day"])

    # Merge the complete date range with the extracted data
    merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"], how='left')
   
    # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]: # Since this is temperature, missing vales cannot be zero (0)
        merged_df[column] = merged_df[column].fillna(np.nan)
    
    # Convert temperature columns to numeric, coerce errors to NaN.
    merged_df['Max_Temperature'] = pd.to_numeric(merged_df['Max_Temperature'], errors='coerce')
    merged_df['Min_Temperature'] = pd.to_numeric(merged_df['Min_Temperature'], errors='coerce')
    merged_df['Avg_Temperature'] = pd.to_numeric(merged_df['Avg_Temperature'], errors='coerce')

    # Sort DataFrame by Year, Month, Day
    merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])

    # Standard deviation and outlier detection
    for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
        # Calculate the standard deviation and mean for the column
        std = merged_df[column].std()
        mean = merged_df[column].mean()

        # # Condition 1: Remove values more than 3 standard deviations from the mean
        # merged_df[column] = merged_df[column].apply(lambda x: np.nan if abs(x - mean) > 3 * std else x)

        # Condition 2: Detect sharp transitions between days (e.g., +4sd to -4sd)
        # Calculate the standard deviation differences for each day relative to the mean
        merged_df['std_diff'] = (merged_df[column] - mean) / std
        for i in range(1, len(merged_df) - 1):  # Avoid the first and last rows to prevent boundary issues
            prev_std_diff = merged_df.loc[i - 1, 'std_diff'] if not pd.isna(merged_df.loc[i - 1, 'std_diff']) else 0   # std deviation difference of previous day
            curr_std_diff = merged_df.loc[i, 'std_diff'] # std deviation difference of current day
            next_std_diff = merged_df.loc[i + 1, 'std_diff'] if not pd.isna(merged_df.loc[i + 1, 'std_diff']) else 0 # std deviation difference of following day

            if not pd.isna(curr_std_diff):
                # Detect sharp opposite changes (e.g., large negative difference to large positive difference or vice versa)
                if (prev_std_diff < -4 and curr_std_diff > 4 and next_std_diff < -4) or \
                    (prev_std_diff > 4 and curr_std_diff < -4 and next_std_diff > 4):
                    merged_df.loc[i, column] = np.nan  # Mark the current value as an outlier

        # Drop the temporary 'std_diff' column
        merged_df.drop(columns=['std_diff'], inplace=True)

    # # List to hold all data
    # data = [] # All the variables

    # total_transcribed_cells = 0

    # # Rows to exclude. Adjust these according to your specific sheet
    # excluded_rows = [1, 2, 3, 9, 10, 16, 17, 23, 24, 30, 31, 37, 38, 45, 46, 47, 48] # These include titles or 5/6 day totals and averages.

    # # Iterate over all files in the folder
    # # Extract year and month from filename
    # year, month = extract_date_from_filename(input_folder_path)
    # if year and month:
    #     workbook = openpyxl.load_workbook(input_folder_path)
    #     worksheet = workbook.active

    #     # Extract data from rows and columns, excluding specific rows.  
    #     for row_num in range(4, worksheet.max_row + 1): #Here this represents Max, Min and Average Temperatures
    #         if row_num not in excluded_rows: 
    #             day_cell = worksheet.cell(row=row_num, column=2)  # Assuming the day is in the first column
    #             max_temperature_cell = worksheet.cell(row=row_num, column=4)  # Column for Max Temperature
    #             min_temperature_cell = worksheet.cell(row=row_num, column=5)  # Column for Min Temperature
    #             average_temperature_cell = worksheet.cell(row=row_num, column=6)  # Column for Avg Temperature
                
    #             # Count the total transcribed values, including even the non-confirmed (GREEN) ones
    #             total_transcribed_cells += sum(cell.value is not None for cell in [max_temperature_cell, min_temperature_cell, average_temperature_cell])

    #             if day_cell.value :
    #                 day = int(day_cell.value)
    #                 max_temperature = max_temperature_cell.value if is_highlighted_green(max_temperature_cell, 'FF6DCD57') else 'NaN'
    #                 min_temperature = min_temperature_cell.value if is_highlighted_green(min_temperature_cell, 'FF6DCD57') else 'NaN'
    #                 average_temperature = average_temperature_cell.value if is_highlighted_green(average_temperature_cell, 'FF6DCD57') else 'NaN'

    #                 data.append([year, month, day, max_temperature, min_temperature, average_temperature])

    # # Create a DataFrame from the data
    # df = pd.DataFrame(data, columns=["Year", "Month", "Day", "Max_Temperature", "Min_Temperature", "Avg_Temperature"])

    # # Generate a complete date range for each year and month combination
    # years_months = df[['Year', 'Month']].drop_duplicates()
    # complete_data = []

    # for _, row in years_months.iterrows():
    #     year = row['Year']
    #     month = row['Month']
    #     num_days = pd.Period(f'{year}-{month}').days_in_month
    #     for day in range(1, num_days + 1):
    #         complete_data.append([year, month, day])

    # complete_df = pd.DataFrame(complete_data, columns=["Year", "Month", "Day"])

    # # Merge the complete date range with the extracted data
    # merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"])

    # # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    # # For combined sheet
    # merged_df['Max_Temperature'] = merged_df['Max_Temperature'].fillna('NaN')  # Since this is temperature, missing vales cannot be zero (0)
    # merged_df['Min_Temperature'] = merged_df['Min_Temperature'].fillna('NaN')
    # merged_df['Avg_Temperature'] = merged_df['Avg_Temperature'].fillna('NaN')
    
    # # Sort DataFrame by Year, Month, Day
    # merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])

    print(f"Total automatically transcribed Cells: {total_transcribed_cells}")

    return merged_df


def select_and_convert_manually_transcribed_data(filepath, date_column, columns_to_check, header_rows, multi_day_totals, multi_day_averages, excluded_rows, additional_excluded_rows, final_totals_rows):

    ''' Selects the confirmed data (from the quality control) and converts the excel file to a format ready to be converted to the Station Exchange Format
    Parameters
    --------------   
    filepath: path of manually transcribed data sheet (file)

    Returns
    -------------- 
    Excel files with selected (confirmed) transcribed data -> in new format 
    Timeseries plot of max, min and avg temperature (confirmed) at particular station
    
    '''

    # Extract year and month from the filename
    filename = os.path.basename(filepath)  # Get the filename from the full file path
    year, month = extract_date_from_filename(filename)  # Assuming `extract_date_from_filename` is a utility that parses the year/month from the filename

    if year and month:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active

    if not (year and month):
        raise ValueError(f"Could not extract year and month from filename: {filename}")


    # List to hold all data
    data = [] # All the variables

    # Lists to hold all data for each temperature type
    data_max = []
    data_min = []
    data_avg = []

    # Rows to exclude. Adjust these according to your specific sheet
    if multi_day_totals and not multi_day_averages:
        excluded_rows =  excluded_rows # These include titles and multi-day (e.g. 5/6) day totals
    if multi_day_totals and multi_day_averages:
        excluded_rows =  excluded_rows + additional_excluded_rows # including both multi day totals and averages
    if not multi_day_totals:
        excluded_rows = final_totals_rows # Exlude only the final totals

    # Convert the day column letter and temperature columns to numeric indices
    date_column_idx = ord(date_column) - ord('A') + 1  # Convert 'B' -> 2  # Date
    column_indices = [ord(col.strip()) - ord('A') + 1 for col in columns_to_check] # Max, min and average temperatures 

    # Now `column_indices` will contain [4, 5, 6] for 'D', 'E', 'F'
    max_temp_column_idx = column_indices[0]  # 'D' column index -> Maximum temperature
    min_temp_column_idx = column_indices[1]  # 'E' column index -> Minimum temperature
    avg_temp_column_idx = column_indices[2]  # 'F' column index -> Avergae temperature

    total_transcribed_cells = 0 

    # Extract data from rows and columns, excluding specific rows.  
    for row_num in range(header_rows+1, worksheet.max_row+1): #Here this represents Max, Min and Average Temperatures
        if row_num not in excluded_rows: 
            day_cell = worksheet.cell(row=row_num, column=date_column_idx)  # Assuming the day is in the first column
            max_temperature_cell = worksheet.cell(row=row_num, column=max_temp_column_idx)  # Column for Max Temperature
            min_temperature_cell = worksheet.cell(row=row_num, column=min_temp_column_idx)  # Column for Min Temperature
            average_temperature_cell = worksheet.cell(row=row_num, column=avg_temp_column_idx)  # Column for Avg Temperature

            # Count the total manually transcribed values
            total_transcribed_cells += sum(cell.value is not None for cell in [max_temperature_cell, min_temperature_cell, average_temperature_cell])

            if day_cell.value is not None and day_cell.value.isdigit():
                day = int(day_cell.value)
                max_temperature = float(max_temperature_cell.value) if max_temperature_cell.value is not None else 'NaN'
                min_temperature = float(min_temperature_cell.value) if min_temperature_cell.value is not None else 'NaN'
                average_temperature = float(average_temperature_cell.value) if average_temperature_cell.value is not None else 'NaN'
                
                data.append([year, month, day, max_temperature, min_temperature, average_temperature])

                data_max.append([year, month, day, max_temperature])
                data_min.append([year, month, day, min_temperature])
                data_avg.append([year, month, day, average_temperature])


    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=["Year", "Month", "Day", "Max_Temperature", "Min_Temperature", "Avg_Temperature"])


    # Generate a complete date range for each year and month combination
    years_months = df[['Year', 'Month']].drop_duplicates()
    complete_data = []

    for _, row in years_months.iterrows():
        year = row['Year']
        month = row['Month']
        num_days = pd.Period(f'{year}-{month}').days_in_month
        for day in range(1, num_days + 1):
            complete_data.append([year, month, day])

    complete_df = pd.DataFrame(complete_data, columns=["Year", "Month", "Day"])

    # Merge the complete date range with the extracted data
    merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"], how='left')
   
    # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]: # Since this is temperature, missing vales cannot be zero (0)
        merged_df[column] = merged_df[column].fillna(np.nan)
    
    # # Convert temperature columns to numeric, coerce errors to NaN.
    # merged_df['Max_Temperature'] = pd.to_numeric(merged_df['Max_Temperature'], errors='coerce')
    # merged_df['Min_Temperature'] = pd.to_numeric(merged_df['Min_Temperature'], errors='coerce')
    # merged_df['Avg_Temperature'] = pd.to_numeric(merged_df['Avg_Temperature'], errors='coerce')

    # Sort DataFrame by Year, Month, Day
    merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])
    # # List to hold all data
    # data = [] # All the variables

    # # Rows to exclude. Adjust these according to your specific sheet
    # excluded_rows = [1, 2, 3, 9, 10, 16, 17, 23, 24, 30, 31, 37, 38, 45, 46, 47, 48] # These include titles or 5/6 day totals and averages.

    # # Iterate over all files in the folder
    # # Extract year and month from filename
    # year, month = extract_date_from_filename(input_folder_path)
    # if year and month:
    #     workbook = openpyxl.load_workbook(input_folder_path)
    #     worksheet = workbook.active

    #     # Extract data from rows and columns, excluding specific rows.  
    #     for row_num in range(4, worksheet.max_row + 1): #Here this represents Max, Min and Average Temperatures
    #         if row_num not in excluded_rows: 
    #             day_cell = worksheet.cell(row=row_num, column=2)  # Assuming the day is in the first column
    #             max_temperature_cell = worksheet.cell(row=row_num, column=4)  # Column for Max Temperature
    #             min_temperature_cell = worksheet.cell(row=row_num, column=5)  # Column for Min Temperature
    #             average_temperature_cell = worksheet.cell(row=row_num, column=6)  # Column for Avg Temperature


    #             if day_cell.value :
    #                 day = int(day_cell.value)
    #                 max_temperature = max_temperature_cell.value 
    #                 min_temperature = min_temperature_cell.value 
    #                 average_temperature = average_temperature_cell.value 
                    
    #                 data.append([year, month, day, max_temperature, min_temperature, average_temperature])

    # # Create a DataFrame from the data
    # df = pd.DataFrame(data, columns=["Year", "Month", "Day", "Max_Temperature", "Min_Temperature", "Avg_Temperature"])

    # # Generate a complete date range for each year and month combination
    # years_months = df[['Year', 'Month']].drop_duplicates()
    # complete_data = []

    # for _, row in years_months.iterrows():
    #     year = row['Year']
    #     month = row['Month']
    #     num_days = pd.Period(f'{year}-{month}').days_in_month
    #     for day in range(1, num_days + 1):
    #         complete_data.append([year, month, day])

    # complete_df = pd.DataFrame(complete_data, columns=["Year", "Month", "Day"])

    # # Merge the complete date range with the extracted data
    # merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"])

    # # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    # # For combined sheet
    # merged_df['Max_Temperature'] = merged_df['Max_Temperature'].fillna('NaN')  # Since this is temperature, missing vales cannot be zero (0)
    # merged_df['Min_Temperature'] = merged_df['Min_Temperature'].fillna('NaN')
    # merged_df['Avg_Temperature'] = merged_df['Avg_Temperature'].fillna('NaN')
    
    # # Sort DataFrame by Year, Month, Day
    # merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])

    print(f"Total manually transcribed Cells: {total_transcribed_cells}")

    return merged_df

# def compare_dataframes(df1, df2, uncertainty_margin):
#     """
#     Compare temperature data between two dataframes and calculate the accuracy percentage.

#     This function compares temperature values from two dataframes, typically one containing manually transcribed data and the other containing post-processed data. 
#     It calculates how closely the values match within a specified uncertainty margin, providing an accuracy percentage as a result.

#     Parameters
#     --------------
#     df1: pandas.DataFrame
#         The first dataframe, usually containing manually transcribed data. It must include columns for 'Year', 'Month', 'Day', and temperature values ('Max_Temperature', 'Min_Temperature', 'Avg_Temperature').
#     df2: pandas.DataFrame
#         The second dataframe, typically containing post-processed data. It should have the same structure and columns as df1.
#     uncertainty_margin: float, optional
#         The allowable difference between values in the two dataframes for them to be considered a match. The default is 0.2.

#     Returns
#     --------------
#     float
#         The accuracy percentage, indicating the proportion of cells in the two dataframes that match within the specified uncertainty margin.
#     """
    
#     total_highlighted_cells = 0
#     accurate_matches = 0

#     # Ensure the dataframes are aligned on the same date
#     merged_df = pd.merge(df1, df2, on=['Year', 'Month', 'Day'], suffixes=('_manual', '_post'))

#     # Convert temperature columns to numeric, coercing errors to NaN
#     for col in ['Max_Temperature', 'Min_Temperature', 'Avg_Temperature']:
#         merged_df[col + '_manual'] = pd.to_numeric(merged_df[col + '_manual'], errors='coerce')
#         merged_df[col + '_post'] = pd.to_numeric(merged_df[col + '_post'], errors='coerce')

#     # Iterate through the relevant columns to compare values
#     for col in ['Max_Temperature', 'Min_Temperature', 'Avg_Temperature']:
#         col_manual = col + '_manual'
#         col_post = col + '_post'

#         for i in range(len(merged_df)):
#             if not pd.isna(merged_df[col_manual].iloc[i]) and not pd.isna(merged_df[col_post].iloc[i]):
#                 total_highlighted_cells += 1
#                 if abs(merged_df[col_manual].iloc[i] - merged_df[col_post].iloc[i]) <= uncertainty_margin:
#                     accurate_matches += 1

#     if total_highlighted_cells == 0:
#         accuracy_percentage = 0.0
#     else:
#         accuracy_percentage = (accurate_matches / total_highlighted_cells) * 100

    
#     print(f"Total Highlighted Cells: {total_highlighted_cells}")
#     print(f"Accuracy Percentage: {accuracy_percentage:.2f}%")

#     return accuracy_percentage


def compare_dataframes(df1, df2, uncertainty_margin): 
    """
    Compare temperature data between two dataframes and calculate the accuracy percentage and mean absolute error (MAE).

    This function compares temperature values from two dataframes, typically one containing manually transcribed data 
    and the other containing post-processed data. It calculates how closely the values match within a specified 
    uncertainty margin, providing an accuracy percentage as a result, and also calculates the mean absolute error.

    Parameters
    --------------
    df1: pandas.DataFrame
        The first dataframe, usually containing manually transcribed data. It must include columns for 'Year', 'Month', 'Day', 
        and temperature values ('Max_Temperature', 'Min_Temperature', 'Avg_Temperature').
    df2: pandas.DataFrame
        The second dataframe, typically containing post-processed data. It should have the same structure and columns as df1.
    uncertainty_margin: float, optional
        The allowable difference between values in the two dataframes for them to be considered a match. The default is 0.2.

    Returns
    --------------
    tuple
        A tuple containing:
        - accuracy_percentage: float
        - mae: float (mean absolute error)
    """
    
    total_highlighted_cells = 0
    accurate_matches = 0
    absolute_errors = []  # List to store absolute errors for MAE calculation

    # Ensure the dataframes are aligned on the same date
    merged_df = pd.merge(df1, df2, on=['Year', 'Month', 'Day'], suffixes=('_manual', '_post'))

    # Convert temperature columns to numeric, coercing errors to NaN
    for col in ['Max_Temperature', 'Min_Temperature', 'Avg_Temperature']:
        merged_df[col + '_manual'] = pd.to_numeric(merged_df[col + '_manual'], errors='coerce')
        merged_df[col + '_post'] = pd.to_numeric(merged_df[col + '_post'], errors='coerce')

    # Iterate through the relevant columns to compare values
    for col in ['Max_Temperature', 'Min_Temperature', 'Avg_Temperature']:
        col_manual = col + '_manual'
        col_post = col + '_post'

        for i in range(len(merged_df)):
            if not pd.isna(merged_df[col_manual].iloc[i]) and not pd.isna(merged_df[col_post].iloc[i]):
                total_highlighted_cells += 1
                difference = abs(merged_df[col_manual].iloc[i] - merged_df[col_post].iloc[i])
                
                # Add the absolute difference to the list for MAE calculation
                absolute_errors.append(difference)

                # Check if the difference is within the uncertainty margin
                if difference <= uncertainty_margin:
                    accurate_matches += 1

    # Calculate accuracy percentage
    accuracy_percentage = (accurate_matches / total_highlighted_cells) * 100 if total_highlighted_cells > 0 else 0.0
    
    # Calculate Mean Absolute Error (MAE)
    mae = sum(absolute_errors) / len(absolute_errors) if absolute_errors else 0.0

    print(f"Total Highlighted Cells: {total_highlighted_cells}")
    print(f"Accuracy Percentage: {accuracy_percentage:.2f}%")
    print(f"Mean Absolute Error (MAE): {mae:.2f}")

    return accuracy_percentage, mae




# def plot_comparison(manual_df, post_processed_df, output_folder_path, station, post_file, accuracy_percentage, uncertainty_margin):
#     """
#     Plot a comparison between manually transcribed data and post-processed data, including confidence intervals.

#     This function generates a visual comparison of daily maximum, minimum, and average temperatures between manually transcribed data and post-processed data for a specific station. 
#     The comparison includes plotting the post-processed data with confidence intervals and overlaying the manually transcribed data for validation purposes.

#     Parameters
#     --------------
#     manual_df: pandas.DataFrame
#         The dataframe containing the manually transcribed temperature data. It should include columns for 'Year', 'Month', 'Day', and the respective temperature values.
#     post_processed_df: pandas.DataFrame
#         The dataframe containing the post-processed temperature data. It should include columns for 'Year', 'Month', 'Day', and the respective temperature values.
#     output_folder_path: str
#         The directory path where the generated plot will be saved.
#     station: str
#         The identifier or name of the station, used for labeling the plot and the output filename.
#     accuracy_percentage: float
#         The calculated accuracy percentage of the manually transcribed data compared to the post-processed data, displayed on the plot.

#     Returns
#     --------------
#     None
#         The function generates and saves a plot comparing the two datasets, displaying the accuracy percentage, and does not return any value.

#     """


#     # Ensure Date column exists in both dataframes
#     manual_df['Date'] = pd.to_datetime(manual_df[['Year', 'Month', 'Day']])
#     post_processed_df['Date'] = pd.to_datetime(post_processed_df[['Year', 'Month', 'Day']])
    
#     # Merge data on Date
#     merged_df = pd.merge(post_processed_df, manual_df, on='Date', suffixes=('_post', '_manual'))
    
#     # Create a figure and axis for the plot
#     fig, ax = plt.subplots(figsize=(12, 8))

#     # Plot manually transcribed data as as points only with 'o' markers
#     ax.plot(merged_df['Date'], merged_df['Max_Temperature_manual'], 'o', label='Max (Manually transcribed)', color='red')
#     ax.plot(merged_df['Date'], merged_df['Min_Temperature_manual'], 'o', label='Min (Manually transcribed)', color='blue')
#     ax.plot(merged_df['Date'], merged_df['Avg_Temperature_manual'], 'o', label='Avg (Manually transcribed)', color='orange')

#     # Plot post-QA/QC data using MeteoSaver (Max, Min, Avg) as dotted lines ('--')
#     ax.plot(merged_df['Date'], merged_df['Max_Temperature_post'], '--', label='Max (Automatically transcribed)', color='red')
#     ax.plot(merged_df['Date'], merged_df['Min_Temperature_post'], '--', label='Min (Automatically transcribed)', color='blue')
#     ax.plot(merged_df['Date'], merged_df['Avg_Temperature_post'], '--', label='Avg (Automatically transcribed)', color='orange')
    
#     # Fill uncertainty margins around post-QA/QC data
#     ax.fill_between(merged_df['Date'], merged_df['Max_Temperature_post'] - uncertainty_margin, 
#                     merged_df['Max_Temperature_post'] + uncertainty_margin, color='red', alpha=0.2)
#     ax.fill_between(merged_df['Date'], merged_df['Min_Temperature_post'] - uncertainty_margin, 
#                     merged_df['Min_Temperature_post'] + uncertainty_margin, color='blue', alpha=0.2)
#     ax.fill_between(merged_df['Date'], merged_df['Avg_Temperature_post'] - uncertainty_margin, 
#                     merged_df['Avg_Temperature_post'] + uncertainty_margin, color='orange', alpha=0.2)
    
#     # Set plot labels and title
#     ax.set_xlabel('Date')
#     ax.set_ylabel('Temperature (°C)')
#     ax.set_title(f'Daily Max, Min, and Avg Temperatures at station {station}')

#     # Adjust the y-axis limit to ensure the legend is above the plotted lines
#     ylim = ax.get_ylim()
#     y_max = max(ylim)
#     legend_y_position = y_max + (y_max * 0.1)
#     ax.set_ylim(ylim[0], legend_y_position)

#     # Plot manually transcribed data as points with 'o' markers
#     manual_handle = [
#     Line2D([0], [0], marker='o', color='red', label='Max (Manually transcribed)', linestyle='None'),
#     Line2D([0], [0], marker='o', color='blue', label='Min (Manually transcribed)', linestyle='None'),
#     Line2D([0], [0], marker='o', color='orange', label='Avg (Manually transcribed)', linestyle='None')]

#     # Create custom legend entries for automatically transcribed data + uncertainty
#     custom_handle = [
#     (Line2D([0], [0], linestyle='--', color='red', lw=2), Patch(facecolor='red', alpha=0.2)),
#     (Line2D([0], [0], linestyle='--', color='blue', lw=2), Patch(facecolor='blue', alpha=0.2)),
#     (Line2D([0], [0], linestyle='--', color='orange', lw=2), Patch(facecolor='orange', alpha=0.2))]

#     # Combine the manual transcribed and automatically transcribed legend handles
#     all_handles = custom_handle + manual_handle

#     # Create custom labels
#     labels = [
#         'Max (Automatically transcribed)', 'Min (Automatically transcribed)', 'Avg (Automatically transcribed)', 'Max (Manually transcribed)', 'Min (Manually transcribed)', 'Avg (Manually transcribed)']

#     # Set legend handles with custom lines for each temperature type
#     ax.legend(handles=all_handles, labels=labels, fontsize='small', loc='upper right')
        
#     ax.grid(True)

#     # Add accuracy percentage below the x-axis
#     plt.figtext(0.5, 0.01, f'Accuracy Percentage: {accuracy_percentage:.1f}%', ha='center', fontsize=10, bbox={"facecolor":"orange", "alpha":0.5, "pad":5})
    
#     # Remove the '.xlsx' extension from the post_file to clean the filename
#     cleaned_post_file_name = os.path.splitext(post_file)[0]

#     # Save the plot
#     plot_filename = os.path.join(output_folder_path, f'temperature_comparison_plot_{cleaned_post_file_name}.jpg')
#     plt.savefig(plot_filename, format='jpg')
#     plt.show()



def plot_comparison(manual_df, post_processed_df, output_folder_path, station, post_file, accuracy_percentage_and_mean_absolute_error, uncertainty_margin):
    """
    Plot a comparison between manually transcribed data and post-processed data, including confidence intervals.

    This function generates a visual comparison of daily maximum, minimum, and average temperatures between manually transcribed data and post-processed data for a specific station. 
    The comparison includes plotting the post-processed data with confidence intervals and overlaying the manually transcribed data for validation purposes.

    Parameters
    --------------
    manual_df: pandas.DataFrame
        The dataframe containing the manually transcribed temperature data. It should include columns for 'Year', 'Month', 'Day', and the respective temperature values.
    post_processed_df: pandas.DataFrame
        The dataframe containing the post-processed temperature data. It should include columns for 'Year', 'Month', 'Day', and the respective temperature values.
    output_folder_path: str
        The directory path where the generated plot will be saved.
    station: str
        The identifier or name of the station, used for labeling the plot and the output filename.
    accuracy_percentage: float
        The calculated accuracy percentage of the manually transcribed data compared to the post-processed data, displayed on the plot.

    Returns
    --------------
    None
        The function generates and saves a plot comparing the two datasets, displaying the accuracy percentage, and does not return any value.

    """


    # Ensure Date column exists in both dataframes
    manual_df['Date'] = pd.to_datetime(manual_df[['Year', 'Month', 'Day']])
    post_processed_df['Date'] = pd.to_datetime(post_processed_df[['Year', 'Month', 'Day']])
    
    # Merge data on Date
    merged_df = pd.merge(post_processed_df, manual_df, on='Date', suffixes=('_post', '_manual'))
    
    # Create a figure and axis for the plot
    fig, ax = plt.subplots(figsize=(12, 8))

    # Plot post-QA/QC data using MeteoSaver (Max, Min, Avg) as dotted lines ('--')
    ax.plot(merged_df['Date'], merged_df['Max_Temperature_post'], 'o', label='Max (Automatically transcribed)', color='red')
    ax.plot(merged_df['Date'], merged_df['Min_Temperature_post'], 'o', label='Min (Automatically transcribed)', color='blue')
    ax.plot(merged_df['Date'], merged_df['Avg_Temperature_post'], 'o', label='Avg (Automatically transcribed)', color='orange')


    # Fill uncertainty margins around post-QA/QC data
    ax.fill_between(merged_df['Date'], merged_df['Max_Temperature_manual'] - uncertainty_margin, 
                    merged_df['Max_Temperature_manual'] + uncertainty_margin, color='red', alpha=0.2)
    ax.fill_between(merged_df['Date'], merged_df['Min_Temperature_manual'] - uncertainty_margin, 
                    merged_df['Min_Temperature_manual'] + uncertainty_margin, color='blue', alpha=0.2)
    ax.fill_between(merged_df['Date'], merged_df['Avg_Temperature_manual'] - uncertainty_margin, 
                    merged_df['Avg_Temperature_manual'] + uncertainty_margin, color='orange', alpha=0.2)
    
    # Set plot labels and title
    ax.set_xlabel('Date', fontsize=17, labelpad=15)
    ax.set_ylabel('Temperature (°C)', fontsize=17, labelpad=15)
    ax.set_title(f'Daily Max, Min, and Avg Temperatures at station {station}')

    # Adjust tick label font size
    ax.tick_params(axis='x', labelsize=17)  # Set font size for x-axis tick labels
    ax.tick_params(axis='y', labelsize=17)  # Set font size for y-axis tick labels

    # Adjust the y-axis limit to ensure the legend is above the plotted lines
    ylim = ax.get_ylim()
    y_max = max(ylim)
    legend_y_position = y_max + (y_max * 0.1)
    ax.set_ylim(ylim[0], legend_y_position)

    # Format the date labels to show at intervals of 5 days
    # Locator for every 5 days
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=5))
    plt.xticks(rotation=45)  # Rotate x-axis tick labels by 45 degrees

    # Add accuracy percentage below the x-axis
    plt.figtext(0.70, 0.92, f'Accuracy Percentage: {accuracy_percentage_and_mean_absolute_error[0]:.1f}%', ha='left', fontsize=15, bbox={"facecolor":"orange", "alpha":0.5, "pad":5})
    plt.figtext(0.70, 0.87, f'Mean Absolute Error: {accuracy_percentage_and_mean_absolute_error[1]:.1f}', ha='left', fontsize=15, bbox={"facecolor":"orange", "alpha":0.5, "pad":5})

    # # LEGEND
    # # Plot manually transcribed data as points with 'o' markers
    # manual_handle = [
    # Line2D([0], [0], marker='o', color='red', label='Max (Automatically transcribed)', linestyle='None'),
    # Line2D([0], [0], marker='o', color='blue', label='Min (Automatically transcribed)', linestyle='None'),
    # Line2D([0], [0], marker='o', color='orange', label='Avg (Automatically transcribed)', linestyle='None')]

    # # Create custom legend entries for automatically transcribed data + uncertainty
    # custom_handle = [
    # ( Patch(facecolor='red', alpha=0.2)),
    # (Patch(facecolor='blue', alpha=0.2)),
    # (Patch(facecolor='orange', alpha=0.2))]

    # # Combine the manual transcribed and automatically transcribed legend handles
    # all_handles =  manual_handle + custom_handle

    # # Create custom labels
    # labels = [
    #     'Max (Automatically transcribed)', 'Min (Automatically transcribed)', 'Avg (Automatically transcribed)', 
    #     'Max (Manually transcribed)', 'Min (Manually transcribed)', 'Avg (Manually transcribed)'
    # ]

    # # Set legend handles with custom lines for each temperature type
    # ax.legend(handles=all_handles, labels=labels, fontsize='small', loc='upper right')

    # Add the legend below the plot
    # ax.legend(handles=all_handles, labels=labels, loc='upper center', bbox_to_anchor=(0.5, -0.40), ncol=2, fontsize=14)
        
    ax.grid(True)

    # # Add accuracy percentage below the x-axis
    # plt.figtext(0.40, 0.01, f'Accuracy Percentage: {accuracy_percentage_and_mean_absolute_error[0]:.1f}%', ha='center', fontsize=10, bbox={"facecolor":"orange", "alpha":0.5, "pad":5})
    # plt.figtext(0.60, 0.01, f'Mean Absolute Error: {accuracy_percentage_and_mean_absolute_error[1]:.1f}', ha='center', fontsize=10, bbox={"facecolor":"orange", "alpha":0.5, "pad":5})
    
    # Remove the '.xlsx' extension from the post_file to clean the filename
    cleaned_post_file_name = os.path.splitext(post_file)[0]

    plt.tight_layout()
    # Save the plot
    plot_filename = os.path.join(output_folder_path, f'temperature_comparison_plot_{cleaned_post_file_name}.jpg')
    plt.savefig(plot_filename, format='jpg')
    plt.show()



def validate(manually_transcribed_data_dir_station, postprocessed_data_dir_station, output_folder_path, station, date_column, columns_to_check, header_rows, multi_day_totals, multi_day_averages, excluded_rows, additional_excluded_rows, final_totals_rows, uncertainty_margin):
    """
    Main function to validate and compare manually transcribed data with post-processed data, and generate corresponding plots.

    This function compares manually transcribed temperature data with post-processed data files to ensure accuracy. 
    It identifies matching files between the two datasets, processes them, and calculates the accuracy percentage. 
    Finally, it generates comparison plots for each file pair.

    Parameters
    --------------
    manually_transcribed_data_dir: str
        The directory path containing manually transcribed data files. These files should be named with the suffix '_manually_entered_temperatures'.
    postprocessed_data_dir_station: str
        The directory path containing post-processed data files. These files should be named with the suffix '_post_QA_QC'.
    output_folder_path: str
        The directory path where the output plots will be saved.
    station: str
        The station identifier or name used for labeling the output files.

    Returns
    --------------
    None
        The function does not return any value. Instead, it processes the data, compares the files, and plots the two files timeseries

    Processing Steps
    --------------
    1. **File Matching:** The function identifies pairs of files from the manually transcribed and post-processed directories that correspond to each other based on their base names.
    2. **Data Loading:** The corresponding data from the identified file pairs are loaded into dataframes.
    3. **Data Conversion:** Temperature columns in the post-processed data are converted to numeric types, with non-convertible values coerced to NaN.
    4. **Accuracy Calculation:** The manually transcribed data is compared with the post-processed data to calculate an accuracy percentage.
    5. **Plot Generation:** Comparison plots are generated, providing a visual assessment of the data validation.
    """

    manually_transcribed_files = os.listdir(manually_transcribed_data_dir_station)
    postprocessed_files = os.listdir(postprocessed_data_dir_station)
    manual_files = [f for f in manually_transcribed_files if 'manually_entered' in f]
    post_processed_files = [f for f in postprocessed_files if 'post_QA_QC' in f]

    # Create a dictionary to map manual files to post-processed files
    file_pairs = {}
    for manual_file in manual_files:
        base_name = manual_file.replace('_manually_entered_temperatures', '').replace('.xlsx', '')
        corresponding_post_file = f'{base_name}_post_QA_QC.xlsx'
        if corresponding_post_file in post_processed_files:
            file_pairs[manual_file] = corresponding_post_file
        else:
            print(f"Post-processed file not found for {manual_file}")

    if not file_pairs:
        print("No matching files found.")
        return

    # Process each file pair
    for manual_file, post_file in file_pairs.items():
        print(f"Processing pair: {manual_file} and {post_file}")
        manual_filepath = os.path.join(manually_transcribed_data_dir_station, manual_file)
        print(manual_filepath)
        post_processed_filepath = os.path.join(postprocessed_data_dir_station, post_file)

        # Load data
        # Manually transcribed data
        manual_df = select_and_convert_manually_transcribed_data(manual_filepath, date_column, columns_to_check, header_rows, multi_day_totals, multi_day_averages, excluded_rows, additional_excluded_rows, final_totals_rows)
        # Automatically transcribed data
        post_processed_df = select_and_convert_postprocessed_data(post_processed_filepath, date_column, columns_to_check, header_rows, multi_day_totals, multi_day_averages, excluded_rows, additional_excluded_rows, final_totals_rows)

        # # Convert temperature columns to numeric, coerce errors to NaN
        # manual_df['Max_Temperature'] = pd.to_numeric(manual_df['Max_Temperature'], errors='coerce')
        # manual_df['Min_Temperature'] = pd.to_numeric(manual_df['Max_Temperature'], errors='coerce')
        # manual_df['Avg_Temperature'] = pd.to_numeric(manual_df['Max_Temperature'], errors='coerce')

        # Convert temperature columns to numeric, coerce errors to NaN for post-processed data
        post_processed_df['Max_Temperature'] = pd.to_numeric(post_processed_df['Max_Temperature'], errors='coerce')
        post_processed_df['Min_Temperature'] = pd.to_numeric(post_processed_df['Min_Temperature'], errors='coerce')
        post_processed_df['Avg_Temperature'] = pd.to_numeric(post_processed_df['Avg_Temperature'], errors='coerce')

        # Check 
        print(manual_df[['Year', 'Month', 'Day']].head())
        print(post_processed_df[['Year', 'Month', 'Day']].head())

        # Compare the dataframes and get accuracy percentage
        accuracy_percentage_and_mean_absolute_error = compare_dataframes(manual_df, post_processed_df, uncertainty_margin)

        # Plot the comparison
        plot_comparison(manual_df, post_processed_df, output_folder_path, station, post_file, accuracy_percentage_and_mean_absolute_error, uncertainty_margin)