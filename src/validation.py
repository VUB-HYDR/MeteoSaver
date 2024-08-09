import os
import openpyxl
import pandas as pd
import re
import matplotlib.pyplot as plt

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

# Create a function to plot with dashed lines for missing data
def plot_with_missing(ax, series, label, color):
    # Plot the main line
    ax.plot(series.index, series, label=label, color=color, linestyle='-')
    
    # Create a mask for missing values
    is_nan = series.isna()
    
    # Find start and end points of missing data segments
    missing_segments = []
    start = None
    for i in range(len(is_nan)):
        if is_nan.iloc[i] and start is None:
            start = i
        elif not is_nan.iloc[i] and start is not None:
            missing_segments.append((start, i))
            start = None
    if start is not None:
        missing_segments.append((start, len(is_nan)))
    
    # Plot dashed lines for missing data
    for start, end in missing_segments:
        if start > 0 and end < len(series):
            ax.plot(series.index[start-1:end+1], series[start-1:end+1], linestyle='--', color=color)


def select_and_convert_postprocessed_data(input_folder_path):

    ''' Selects the confirmed data (from the quality control) and converts the excel file to a format ready to be converted to the Station Exchange Format
    Parameters
    --------------   
    input_folder_path: path of postprocessed files
    output_folder_path: output path for the selected data
    station: station number

    Returns
    -------------- 
    Excel files with selected (confirmed) transcribed data -> in new format 
    Timeseries plot of max, min and avg temperature (confirmed) at particular station
    
    '''

    # List to hold all data
    data = [] # All the variables

    total_transcribed_cells = 0

    # Rows to exclude. Adjust these according to your specific sheet
    excluded_rows = [1, 2, 3, 9, 10, 16, 17, 23, 24, 30, 31, 37, 38, 45, 46, 47, 48] # These include titles or 5/6 day totals and averages.

    # Iterate over all files in the folder
    # Extract year and month from filename
    year, month = extract_date_from_filename(input_folder_path)
    if year and month:
        workbook = openpyxl.load_workbook(input_folder_path)
        worksheet = workbook.active

        # Extract data from rows and columns, excluding specific rows.  
        for row_num in range(4, worksheet.max_row + 1): #Here this represents Max, Min and Average Temperatures
            if row_num not in excluded_rows: 
                day_cell = worksheet.cell(row=row_num, column=2)  # Assuming the day is in the first column
                max_temperature_cell = worksheet.cell(row=row_num, column=4)  # Column for Max Temperature
                min_temperature_cell = worksheet.cell(row=row_num, column=5)  # Column for Min Temperature
                average_temperature_cell = worksheet.cell(row=row_num, column=6)  # Column for Avg Temperature
                
                # Count the total transcribed values, including even the non-confirmed (GREEN) ones
                total_transcribed_cells += sum(cell.value is not None for cell in [max_temperature_cell, min_temperature_cell, average_temperature_cell])

                if day_cell.value :
                    day = int(day_cell.value)
                    max_temperature = max_temperature_cell.value if is_highlighted_green(max_temperature_cell, 'FF6DCD57') else 'NaN'
                    min_temperature = min_temperature_cell.value if is_highlighted_green(min_temperature_cell, 'FF6DCD57') else 'NaN'
                    average_temperature = average_temperature_cell.value if is_highlighted_green(average_temperature_cell, 'FF6DCD57') else 'NaN'

                    data.append([year, month, day, max_temperature, min_temperature, average_temperature])

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
    merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"])

    # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    # For combined sheet
    merged_df['Max_Temperature'] = merged_df['Max_Temperature'].fillna('NaN')  # Since this is temperature, missing vales cannot be zero (0)
    merged_df['Min_Temperature'] = merged_df['Min_Temperature'].fillna('NaN')
    merged_df['Avg_Temperature'] = merged_df['Avg_Temperature'].fillna('NaN')
    
    # Sort DataFrame by Year, Month, Day
    merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])

    print(f"Total Transcribed Cells: {total_transcribed_cells}")

    return merged_df


def select_and_convert_manually_transcribed_data(input_folder_path):

    ''' Selects the confirmed data (from the quality control) and converts the excel file to a format ready to be converted to the Station Exchange Format
    Parameters
    --------------   
    input_folder_path: path of postprocessed files
    output_folder_path: output path for the selected data
    station: station number

    Returns
    -------------- 
    Excel files with selected (confirmed) transcribed data -> in new format 
    Timeseries plot of max, min and avg temperature (confirmed) at particular station
    
    '''

    # List to hold all data
    data = [] # All the variables

    # Rows to exclude. Adjust these according to your specific sheet
    excluded_rows = [1, 2, 3, 9, 10, 16, 17, 23, 24, 30, 31, 37, 38, 45, 46, 47, 48] # These include titles or 5/6 day totals and averages.

    # Iterate over all files in the folder
    # Extract year and month from filename
    year, month = extract_date_from_filename(input_folder_path)
    if year and month:
        workbook = openpyxl.load_workbook(input_folder_path)
        worksheet = workbook.active

        # Extract data from rows and columns, excluding specific rows.  
        for row_num in range(4, worksheet.max_row + 1): #Here this represents Max, Min and Average Temperatures
            if row_num not in excluded_rows: 
                day_cell = worksheet.cell(row=row_num, column=2)  # Assuming the day is in the first column
                max_temperature_cell = worksheet.cell(row=row_num, column=4)  # Column for Max Temperature
                min_temperature_cell = worksheet.cell(row=row_num, column=5)  # Column for Min Temperature
                average_temperature_cell = worksheet.cell(row=row_num, column=6)  # Column for Avg Temperature


                if day_cell.value :
                    day = int(day_cell.value)
                    max_temperature = max_temperature_cell.value 
                    min_temperature = min_temperature_cell.value 
                    average_temperature = average_temperature_cell.value 
                    
                    data.append([year, month, day, max_temperature, min_temperature, average_temperature])

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
    merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"])

    # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    # For combined sheet
    merged_df['Max_Temperature'] = merged_df['Max_Temperature'].fillna('NaN')  # Since this is temperature, missing vales cannot be zero (0)
    merged_df['Min_Temperature'] = merged_df['Min_Temperature'].fillna('NaN')
    merged_df['Avg_Temperature'] = merged_df['Avg_Temperature'].fillna('NaN')
    
    # Sort DataFrame by Year, Month, Day
    merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])

    return merged_df

def compare_dataframes(df1, df2, uncertainty_margin=0.2):
    
    total_highlighted_cells = 0
    accurate_matches = 0


    # Ensure the dataframes are aligned on the same date
    merged_df = pd.merge(df1, df2, on=['Year', 'Month', 'Day'], suffixes=('_manual', '_post'))

    # Iterate through the relevant columns to compare values
    for col in ['Max_Temperature', 'Min_Temperature', 'Avg_Temperature']:
        col_manual = col + '_manual'
        col_post = col + '_post'

        for i in range(len(merged_df)):
            if not pd.isna(merged_df[col_manual].iloc[i]) and not pd.isna(merged_df[col_post].iloc[i]):
                total_highlighted_cells += 1
                if abs(merged_df[col_manual].iloc[i] - merged_df[col_post].iloc[i]) <= uncertainty_margin:
                    accurate_matches += 1

    if total_highlighted_cells == 0:
        accuracy_percentage = 0.0
    else:
        accuracy_percentage = (accurate_matches / total_highlighted_cells) * 100

    
    print(f"Total Highlighted Cells: {total_highlighted_cells}")
    print(f"Accuracy Percentage: {accuracy_percentage:.2f}%")

    return accuracy_percentage

def plot_comparison(manual_df, post_processed_df, output_folder_path, station, accuracy_percentage):
    """Plot the comparison of manually entered data with post-processed data."""
    # Ensure Date column exists in both dataframes
    manual_df['Date'] = pd.to_datetime(manual_df[['Year', 'Month', 'Day']])
    post_processed_df['Date'] = pd.to_datetime(post_processed_df[['Year', 'Month', 'Day']])
    
    # Merge data on Date
    merged_df = pd.merge(post_processed_df, manual_df, on='Date', suffixes=('_post', '_manual'))
    
    # Create a figure and axis for the plot
    fig, ax = plt.subplots(figsize=(12, 8))

    # Assuming the standard error is 0.2 for the sake of demonstration
    standard_error = 0.2   # here i used the uncertainity margin allowed in the post processing of the transcribed data
    z = 1.96  # Z-score for 95% confidence interval
    
    # Plot Max Temperature with confidence interval band using post-processed data
    ax.plot(merged_df['Date'], merged_df['Max_Temperature_post'], label='Max (Post-QA/QC data: MeteoSaver)', color='red', marker='o')
    ax.fill_between(merged_df['Date'], merged_df['Max_Temperature_post'] - z * standard_error, merged_df['Max_Temperature_post'] + z * standard_error, color='red', alpha=0.2)
    
    # Plot Min Temperature with confidence interval band using post-processed data
    ax.plot(merged_df['Date'], merged_df['Min_Temperature_post'], label='Min (Post-QA/QC data: MeteoSaver)', color='blue', marker='o')
    ax.fill_between(merged_df['Date'], merged_df['Min_Temperature_post'] - z * standard_error, merged_df['Min_Temperature_post'] + z * standard_error, color='blue', alpha=0.2)
    
    # Plot Avg Temperature with confidence interval band using post-processed data
    ax.plot(merged_df['Date'], merged_df['Avg_Temperature_post'], label='Avg (Post-QA/QC data: MeteoSaver)', color='orange', marker='o')
    ax.fill_between(merged_df['Date'], merged_df['Avg_Temperature_post'] - z * standard_error, merged_df['Avg_Temperature_post'] + z * standard_error, color='orange', alpha=0.2)
    
    # Plot manually entered data as black dotted lines
    ax.plot(merged_df['Date'], merged_df['Max_Temperature_manual'], '--', label='Max (Manually transcribed)', color='red')
    ax.plot(merged_df['Date'], merged_df['Min_Temperature_manual'], '--', label='Min(Manually transcribed)', color ='blue')
    ax.plot(merged_df['Date'], merged_df['Avg_Temperature_manual'], '--', label='Avg (Manually transcribed)', color ='orange')
    
    # Set plot labels and title
    ax.set_xlabel('Date')
    ax.set_ylabel('Temperature (°C)')
    ax.set_title(f'Daily Max, Min, and Avg Temperatures at station {station}')
    #ax.legend(loc='upper right', fontsize='small')
    # Adjust the y-axis limit to ensure the legend is above the plotted lines
    ylim = ax.get_ylim()
    y_max = max(ylim)
    legend_y_position = y_max + (y_max * 0.1)
    ax.set_ylim(ylim[0], legend_y_position)

    # Adjust the legend to be above the plotted lines
    ax.legend(loc='upper right', fontsize='small')
    

    ax.grid(True)

    # Add accuracy percentage below the x-axis
    plt.figtext(0.5, 0.01, f'Accuracy Percentage: {accuracy_percentage:.2f}%', ha='center', fontsize=10, bbox={"facecolor":"orange", "alpha":0.5, "pad":5})
    
    # Save the plot
    plot_filename = os.path.join(output_folder_path, f'temperature_comparison_plot_{station}.jpg')
    plt.savefig(plot_filename, format='jpg')
    plt.show()

def validate(manually_transcribed_data_dir, postprocessed_data_dir_station, output_folder_path, station):
    """Main function to process and plot data."""
    manually_transcribed_files = os.listdir(manually_transcribed_data_dir)
    postprocessed_files = os.listdir(postprocessed_data_dir_station)
    manual_files = [f for f in manually_transcribed_files if 'manually_entered' in f]
    post_processed_files = [f for f in postprocessed_files if 'post_processed' in f]

    # Create a dictionary to map manual files to post-processed files
    file_pairs = {}
    for manual_file in manual_files:
        base_name = manual_file.replace('_manually_entered_temperatures', '').replace('.xlsx', '')
        corresponding_post_file = f'{base_name}_post_processed.xlsx'
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
        manual_filepath = os.path.join(manually_transcribed_data_dir, manual_file)
        post_processed_filepath = os.path.join(postprocessed_data_dir_station, post_file)

        # Load data
        #manual_df = load_data(manual_filepath)
        manual_df = select_and_convert_manually_transcribed_data(manual_filepath)
        #post_processed_df = load_data(post_processed_filepath)
        post_processed_df = select_and_convert_postprocessed_data(post_processed_filepath)

        # Convert temperature columns to numeric, coerce errors to NaN for post-processed data
        post_processed_df['Max_Temperature'] = pd.to_numeric(post_processed_df['Max_Temperature'], errors='coerce')
        post_processed_df['Min_Temperature'] = pd.to_numeric(post_processed_df['Min_Temperature'], errors='coerce')
        post_processed_df['Avg_Temperature'] = pd.to_numeric(post_processed_df['Avg_Temperature'], errors='coerce')

        # Compare the dataframes and get accuracy percentage
        accuracy_percentage = compare_dataframes(manual_df, post_processed_df)

        # Plot the comparison
        plot_comparison(manual_df, post_processed_df, output_folder_path, station, accuracy_percentage)