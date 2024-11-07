import os
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import re
import matplotlib.pyplot as plt
import numpy as np


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
    '''
    Plots a time series with a solid line, highlighting segments with missing data using dashed lines.

    This function plots a time series on the given Axes object (`ax`). The main data is represented by a solid line. 
    Where there are missing values (`NaN`) in the series, dashed lines are plotted to indicate the gaps.

    Parameters
    --------------
    ax: matplotlib.axes.Axes
        The matplotlib Axes object where the series will be plotted.
    series: pandas.Series
        The time series data to plot. The index should represent the x-axis (e.g., dates), and the values represent the y-axis data.
    label: str
        The label for the series, used for the legend.
    color: str
        The color of the plot line.

    Returns
    --------------
    None
        The function plots the series directly onto the provided Axes object and does not return any value.
    '''


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


def dms_to_decimal(dms_str):
    """Convert DMS (degrees, minutes, seconds) string to decimal degrees."""

    if not isinstance(dms_str, str):
        # Return NaN if the input is not a valid string
        return np.nan
    
    # Check if there is a direction (N/S/E/W)
    match = re.match(r'([NSWE])?\s*(\d+)°(\d+)', dms_str)
    if not match:
        raise ValueError(f"Invalid DMS format: {dms_str}")
    
    direction, degrees, minutes = match.groups()
    decimal = int(degrees) + int(minutes) / 60
    
    # Make the decimal negative for S and W directions if direction is specified
    if direction in ['S', 'W']:
        decimal = -decimal
    
    return round(decimal, 4)

def load_station_metadata(file_path, sheet_name='Stations'):
    """Load station metadata from Excel file and convert latitude/longitude to decimal degrees."""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df = df.rename(columns={
        'Station': 'name',
        'Station ID': 'ID',
        'Latitude': 'latitude',
        'Longitude': 'longitude',
        'Altitude': 'altitude'
    })
    
    # Trim whitespace from IDs and ensure they're all strings
    df['ID'] = df['ID'].astype(str).str.strip()

    # Convert latitude and longitude to decimal degrees using dms_to_decimal
    df['latitude'] = df['latitude'].apply(dms_to_decimal)
    df['longitude'] = df['longitude'].apply(dms_to_decimal)
    
    return df[['ID', 'name', 'latitude', 'longitude', 'altitude']]



def convert_to_sef_with_metadata(df, station_info, temp_column, temp_type, source="Institut National pour l’Etude et la Recherche Agronomiques", link="", stat="point", units="C", observer=""):
    """Convert DataFrame to SEF format for a specific temperature type using station metadata."""
    
    # Define SEF headers as a list of strings
    sef_headers = {
        "SEF": "1.0.0",
        "ID": station_info['ID'],
        "Name": station_info['name'],
        "Lat": station_info['latitude'],
        "Lon": station_info['longitude'],
        "Alt": station_info['altitude'],
        "Source": source,
        "Link": link,
        "Vbl": temp_type,
        "Stat": stat,
        "Units": units,
        "Meta": f"Observer={observer}  | QC software = MeteoSaver v1.0 | Note = Transcription software: MeteoSaver v1.0 (https://github.com/VUB-HYDR/MeteoSaver)"
    }

    
    # Prepare the SEF data rows
    sef_df = pd.DataFrame({
        "Year": df["Year"],
        "Month": df["Month"],
        "Day": df["Day"],
        "Hour": [0] * len(df),   # Assuming midnight for simplicity
        "Minute": [0] * len(df),
        "Period": [0] * len(df),
        "Value": df[temp_column].fillna("NaN"),
        "Meta": [""] * len(df)   # Placeholder for any additional meta information
    })
    
    # Define the correct column order for SEF data
    sef_column_order = ["Year", "Month", "Day", "Hour", "Minute", "Period", "Value", "Meta"]
    sef_df = sef_df[sef_column_order]

    return sef_headers, sef_df



def data_formatting(input_folder_path, output_folder_path, metadata_file_path, station, date_column, columns_to_check, header_rows, multi_day_totals, multi_day_averages, excluded_rows, additional_excluded_rows, final_totals_rows, uncertainty_margin):

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

    output_file = os.path.join(output_folder_path, 'Daily_all_temperatures.xlsx')  # Combined output file with the three variables: Max, Min, and Average Temperature
    output_files = {  # Output files for individual temperature columns
        'Max_Temperature': os.path.join(output_folder_path, 'Daily_max_temperatures.xlsx'),
        'Min_Temperature': os.path.join(output_folder_path, 'Daily_min_temperatures.xlsx'),
        'Avg_Temperature': os.path.join(output_folder_path, 'Daily_mean_temperatures.xlsx')
    }

    # Load station metadata
    station_metadata = load_station_metadata(metadata_file_path)
    
    # Ensure both station ID and station parameter are strings
    station = str(station)  # Convert input station ID to string
    station_metadata['ID'] = station_metadata['ID'].astype(str)  # Ensure IDs in metadata are also strings
    print("Station metadata IDs:", station_metadata['ID'].tolist())
    
    # Filter metadata for the specified station ID
    station_info = station_metadata[station_metadata['ID'] == station]
    if station_info.empty:
        raise ValueError(f"No metadata found for station ID {station}")
    station_info = station_info.iloc[0]  # Convert to a Series

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

    # Iterate over all files in the folder
    for filename in os.listdir(input_folder_path):
        if filename.endswith(".xlsx"):
            # Extract year and month from filename
            year, month = extract_date_from_filename(filename)
            if year and month:
                file_path = os.path.join(input_folder_path, filename)
                workbook = openpyxl.load_workbook(file_path)
                worksheet = workbook.active

                # Extract data from rows and columns, excluding specific rows.  
                for row_num in range(header_rows+1, worksheet.max_row + 1): #Here this represents Max, Min and Average Temperatures
                    if row_num not in excluded_rows: 
                        day_cell = worksheet.cell(row=row_num, column=date_column_idx)  # Assuming the day is in the first column
                        max_temperature_cell = worksheet.cell(row=row_num, column=max_temp_column_idx)  # Column for Max Temperature
                        min_temperature_cell = worksheet.cell(row=row_num, column=min_temp_column_idx)  # Column for Min Temperature
                        average_temperature_cell = worksheet.cell(row=row_num, column=avg_temp_column_idx)  # Column for Avg Temperature


                        if day_cell.value :
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
    merged_df = pd.merge(complete_df, df, on=["Year", "Month", "Day"])
   
    # Fill missing temperatures with a placeholder value (e.g., NaN or a specific value)
    for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]: # Since this is temperature, missing vales cannot be zero (0)
        merged_df[column] = merged_df[column].fillna(np.nan)
    
    # Convert temperature columns to numeric, coerce errors to NaN.
    merged_df['Max_Temperature'] = pd.to_numeric(merged_df['Max_Temperature'], errors='coerce')
    merged_df['Min_Temperature'] = pd.to_numeric(merged_df['Min_Temperature'], errors='coerce')
    merged_df['Avg_Temperature'] = pd.to_numeric(merged_df['Avg_Temperature'], errors='coerce')
    
    # Sort DataFrame by Year, Month, Day
    merged_df = merged_df.sort_values(by=['Year', 'Month', 'Day'])

    # Standard deviation and outlier detection with flagging
    for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
        std = merged_df[column].std()
        mean = merged_df[column].mean()

        merged_df[f"{column}_Flag"] = ""
        
        ## UNCOMMENT BELOW FOR CONDITION 1 IN CASE OF LONG TIME SERIES
        # # Condition 1: Remove values > 3 std deviations from the mean and flag them
        
        # for i in range(len(merged_df)):
        #     if abs(merged_df.loc[i, column] - mean) > 3 * std:
        #         merged_df.loc[i, column] = np.nan
        #         merged_df.loc[i, f"{column}_Flag"] = "Condition 1"  # Flag as Condition 1 (current value as an outlier)

        # Condition 2: Detect and flag sharp transitions between days (e.g., from -4std to +4sd to -4sd)
        # Calculate the standard deviation differences for each day relative to the mean
        merged_df['std_diff'] = (merged_df[column] - mean) / std
        for i in range(1, len(merged_df) - 1): # Avoid the first and last rows to prevent boundary issues
            prev_std_diff = merged_df.loc[i - 1, 'std_diff'] if not pd.isna(merged_df.loc[i - 1, 'std_diff']) else 0 # std deviation difference of previous day
            curr_std_diff = merged_df.loc[i, 'std_diff'] # std deviation difference of current day
            next_std_diff = merged_df.loc[i + 1, 'std_diff'] if not pd.isna(merged_df.loc[i + 1, 'std_diff']) else 0 # std deviation difference of following day

            # Detect sharp opposite changes (e.g., large negative difference to large positive difference or vice versa)
            if not pd.isna(curr_std_diff) and (
                (prev_std_diff < -4 and curr_std_diff > 4 and next_std_diff < -4) or
                (prev_std_diff > 4 and curr_std_diff < -4 and next_std_diff > 4)
            ):
                merged_df.loc[i, column] = np.nan
                merged_df.loc[i, f"{column}_Flag"] = "Condition 2"  # Flag as Condition 2 (current value as an outlier)

        # Drop the temporary std_diff column
        merged_df.drop(columns=['std_diff'], inplace=True)

    # Save flagged data to Excel and apply conditional formatting
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="Data")
        workbook = writer.book
        worksheet = writer.sheets["Data"]

        # Define the dark red fill for flagged cells
        dark_red_fill = PatternFill(start_color="CC3300", end_color="CC3300", fill_type="solid")

        # Apply conditional formatting to only flagged cells
        for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
            flag_column = f"{column}_Flag"
            for row in range(2, len(merged_df) + 2):  # Adjusting for header in Excel
                if merged_df.loc[row - 2, flag_column] in ["Condition 1", "Condition 2"]:
                    cell = worksheet[f"{openpyxl.utils.get_column_letter(merged_df.columns.get_loc(column) + 1)}{row}"]
                    cell.fill = dark_red_fill

    # Clean up flag columns in the DataFrame for further processing, if needed
    merged_df.drop(columns=[f"{col}_Flag" for col in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]], inplace=True)

    # # Save the DataFrame to a new Excel file 
    # #merged_df.to_excel(output_file, index=False)
    # timeseries = merged_df.fillna('NaN')
    # timeseries.to_excel(output_file, index=False)

    # After processing, generate the SEF file
    # Loop over each temperature type and create a SEF file for each
    temperature_columns = {
        "Max_Temperature": "Tx",
        "Min_Temperature": "Tn",
        "Avg_Temperature": "Ta"
    }
    
    
    for temp_column, temp_type in temperature_columns.items():
        # Filter data for the specific temperature type
        timeseries_df = merged_df[['Year', 'Month', 'Day', temp_column]].fillna('NaN')
        timeseries_df = timeseries_df.rename(columns={temp_column: "Value"})

        # Convert to SEF format with headers using the function
        sef_headers, sef_df = convert_to_sef_with_metadata(
            df=timeseries_df,
            station_info=station_info,
            temp_column="Value",    # Pass the renamed column "Value"
            temp_type=temp_type      # Pass the type (e.g., Tx, Tn, Ta) for SEF header
        )

        # Define the output file path
        sef_output_file = os.path.join(output_folder_path, f"SEF_station_{station}_{temp_type}_temperature.tsv")
        
        # Write headers and data to the TSV file
        with open(sef_output_file, 'w') as f:
            # Write each header line with tab separation
            for key, value in sef_headers.items():
                f.write(f"{key}\t{value}\n")
            

            # Write the main SEF data with tab separation and include the header row for data columns
            sef_df.to_csv(f, index=False, sep='\t', header=True)


    # # Save to individual files
    # for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
    #     output_file = output_files[column]
    #     timeseries = merged_df[['Year', 'Month', 'Day', column]].fillna('NaN')
    #     timeseries.to_excel(output_file, index=False)

    
    # # Convert temperature columns to numeric, coerce errors to NaN.
    # merged_df['Max_Temperature'] = pd.to_numeric(merged_df['Max_Temperature'], errors='coerce')
    # merged_df['Min_Temperature'] = pd.to_numeric(merged_df['Min_Temperature'], errors='coerce')
    # merged_df['Avg_Temperature'] = pd.to_numeric(merged_df['Avg_Temperature'], errors='coerce')


    # Plot and save the graph
    merged_df['Date'] = pd.to_datetime(merged_df[['Year', 'Month', 'Day']])
    
  
    #***PLOTTING***
    # ## Plot timeseries: lines with breaks in cases with missing data
    # Do not drop rows with NaN values
    plot_df = merged_df[['Date', 'Max_Temperature', 'Min_Temperature', 'Avg_Temperature']]

    # Set the Date column as the index
    plot_df.set_index('Date', inplace=True)

    # Create a figure and axis for the plot
    fig, ax = plt.subplots(figsize=(10, 6))

    # Plot Max Temperature with confidence interval band using pandas plot function
    plot_df['Max_Temperature'].plot(ax=ax, label='Maximum', color='red')
    ax.fill_between(plot_df.index, plot_df['Max_Temperature'] - uncertainty_margin, plot_df['Max_Temperature'] + uncertainty_margin, color='red', alpha=0.2)

    # Plot Min Temperature with confidence interval band using pandas plot function
    plot_df['Min_Temperature'].plot(ax=ax, label='Minimum', color='blue')
    ax.fill_between(plot_df.index, plot_df['Min_Temperature'] - uncertainty_margin, plot_df['Min_Temperature'] + uncertainty_margin, color='blue', alpha=0.2)

    # Plot Avg Temperature with confidence interval band using pandas plot function
    plot_df['Avg_Temperature'].plot(ax=ax, label='Average', color='orange')
    ax.fill_between(plot_df.index, plot_df['Avg_Temperature'] - uncertainty_margin, plot_df['Avg_Temperature'] + uncertainty_margin, color='orange', alpha=0.2)

    # Set plot labels and title
    ax.set_xlabel('Date')
    ax.set_ylabel('Temperature (°C)')
    ax.set_title('Daily Maximum, Minimum, and Average Temperatures at Station ' + str(station))
    ax.legend()
    ax.grid(True)

    # Save the plot
    plt.savefig(f'{output_folder_path}/temperature_plot.jpg', format='jpg')
    plt.show()
