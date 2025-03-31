import os
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import re
import matplotlib.pyplot as plt
import numpy as np
from sklearn.linear_model import LinearRegression
from scipy.stats import theilslopes


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


def dms_to_decimal(dms_str):
    """
    Convert a DMS (Degrees, Minutes, Seconds) string to decimal degrees.

    This function takes a string representing geographic coordinates in the 
    Degrees-Minutes-Seconds (DMS) format and converts it to a decimal degrees 
    representation. It accounts for directional indicators (N, S, E, W) 
    to determine the sign of the decimal value.

    Parameters
    ----------
    dms_str : str
        The DMS string to be converted. The expected format includes an optional 
        direction (N/S/E/W), followed by degrees (°) and minutes ("), e.g., "N 45°30".

    Returns
    -------
    float
        The converted decimal degrees value, rounded to four decimal places. Returns
        NaN if the input is not a valid DMS string.

    Raises
    ------
    ValueError
        If the input string does not match the expected DMS format.

    
    For example:
    --------
    dms_to_decimal("N 45°30") to 45.5000
    """

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
    """
    Load station metadata from an Excel file and convert geographic coordinates to decimal degrees.

    This function reads station metadata from a specified Excel file and processes it to standardize
    column names, trim whitespace, and convert latitude and longitude values from Degrees-Minutes-Seconds 
    (DMS) format to decimal degrees. It returns a cleaned DataFrame with essential station details.

    Parameters
    ----------
    file_path : str
        The file path to the Excel file containing the station metadata.
    sheet_name : str, optional
        The name of the Excel sheet to read. Defaults to 'Stations'.

    Returns
    -------
    pandas.DataFrame
        A DataFrame containing cleaned and processed station metadata with the following columns:
        - 'ID': Unique identifier for each station (station no.).
        - 'name': Station name.
        - 'latitude': Latitude in decimal degrees.
        - 'longitude': Longitude in decimal degrees.
        - 'altitude': Altitude in meters.
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    df = df.rename(columns={
        'Station': 'name',
        'ID': 'ID',
        'Latitude': 'latitude',
        'Longitude': 'longitude',
        'Altitude': 'altitude'
    })
    
    # Trim whitespace from IDs and ensure they're all strings
    df['ID'] = df['ID'].astype(str).str.strip().str.zfill(3)

    # Convert latitude and longitude to decimal degrees using dms_to_decimal
    df['latitude'] = df['latitude'].apply(dms_to_decimal)
    df['longitude'] = df['longitude'].apply(dms_to_decimal)
    
    return df[['ID', 'name', 'latitude', 'longitude', 'altitude']]



def analyze_temperature_trends_with_linear_regression(output_folder_path, station, station_name):
    """
    Analyzes and plots temperature trends over time using linear regression.

    Parameters
    ----------
    output_folder_path : str
        Path where the processed temperature dataset is stored.
    station : str
        Unique station identifier.

    Returns
    -------
    trend_slopes : dict
        Slopes of the temperature trends for Max, Min, and Avg temperatures.
    """

    # Load the dataset
    file_path = os.path.join(output_folder_path, f"Daily_temperature_and_precipitation_station_{station}.xlsx")

    if not os.path.exists(file_path):
        print(f"Error: File not found - {file_path}")
        return None

    df = pd.read_excel(file_path, sheet_name="Data")

    # Drop rows with missing temperature values
    df_clean = df.dropna(subset=["Max_Temperature", "Min_Temperature", "Avg_Temperature"]).copy()

    # Convert the date components into a single datetime column
    df_clean["Date"] = pd.to_datetime(df_clean[["Year", "Month", "Day"]])

    # Convert the Year to numerical values for regression
    df_clean["Year_Num"] = df_clean["Year"].astype(int)

    # Check if df_clean is empty
    if df_clean.empty:
        print(f"Warning: No valid data at station {station}. Debugging:")
        print(df.head())  # Debugging step
        return None

    # Fit linear regression models for Max, Min, and Avg temperatures
    X = df_clean["Year_Num"].values.reshape(-1, 1)  # Predictor variable (Year)

    models = {}
    predictions = {}

    for temp_type in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
        y = df_clean[temp_type].values  # Response variable (Temperature)
        
        if len(y) == 0:
            print(f"Warning: No valid {temp_type} data at station {station}. Skipping.")
            continue
        
        model = LinearRegression().fit(X, y)
        models[temp_type] = model
        predictions[temp_type] = model.predict(X)

    # Ensure at least one model was trained
    if not models:
        print(f"Error: No valid data for regression at station {station}. Skipping plot.")
        return None

    # Plot the trends
    plt.figure(figsize=(12, 6))
    colors = {"Max_Temperature": "red", "Min_Temperature": "blue", "Avg_Temperature": "orange"}
    trend_colors = {"Max_Temperature": "darkred", "Min_Temperature": "darkblue", "Avg_Temperature": "darkorange"}

    for temp_type, model in models.items():
        plt.scatter(df_clean["Date"], df_clean[temp_type], color=colors[temp_type], alpha=0.5, s=10,
                label=f"{temp_type.replace('_', ' ')}")
        plt.plot(df_clean["Date"], predictions[temp_type], color=trend_colors[temp_type], linewidth=2,label=f"{temp_type.replace('_', ' ')} trend")
        
        # Extract slope value for annotation
        slope = model.coef_[0]

        # Get last date and predicted temperature
        last_date = df_clean["Date"].iloc[-1]
        last_predicted_temp = predictions[temp_type][-1]

        # Annotate slope on the trend line
        plt.text(last_date, last_predicted_temp, f"{slope:.3f} °C/yr", 
                 fontsize=10, color=trend_colors[temp_type], fontweight='bold')

    plt.xlabel("Year")
    plt.ylabel("Temperature (°C)")
    plt.title(f"Temperature Trends at Station: {station_name} - using Linear Regression")
    plt.legend()
    plt.grid(True)

    # Save the plot
    os.makedirs(output_folder_path, exist_ok=True)
    trend_plot_path = os.path.join(output_folder_path, f"temperature_trend_plot_{station}_using_linear_regression.jpg")
    plt.savefig(trend_plot_path, format="jpg", dpi=300)
    print(f"Figure saved at: {trend_plot_path}")

    # plt.show()
    plt.close()  # Close after saving

    return {temp: models[temp].coef_[0] for temp in models}



def analyze_temperature_trends_with_theilsen(output_folder_path, station, station_name):
    """
    Analyzes and plots temperature trends over time using both linear regression and Theil–Sen estimator.

    Parameters
    ----------
    output_folder_path : str
        Path where the processed temperature dataset is stored.
    station : str
        Unique station identifier.
    station_name : str
        Readable station name for plotting.

    Returns
    -------
    trend_slopes : dict
        Slopes of the temperature trends for Max, Min, and Avg temperatures using both estimators.
    """
    file_path = os.path.join(output_folder_path, f"Daily_temperature_and_precipitation_station_{station}.xlsx")

    if not os.path.exists(file_path):
        print(f"Error: File not found - {file_path}")
        return None

    df = pd.read_excel(file_path, sheet_name="Data")
    df_clean = df.dropna(subset=["Max_Temperature", "Min_Temperature", "Avg_Temperature"]).copy()
    df_clean["Date"] = pd.to_datetime(df_clean[["Year", "Month", "Day"]])
    df_clean["Year_Num"] = df_clean["Year"].astype(int)

    if df_clean.empty:
        print(f"Warning: No valid data at station {station}.")
        return None

    colors = {"Max_Temperature": "red", "Min_Temperature": "blue", "Avg_Temperature": "orange"}
    trend_colors = {"Max_Temperature": "darkred", "Min_Temperature": "darkblue", "Avg_Temperature": "darkorange"}

    trend_slopes = {}

    plt.figure(figsize=(12, 6))

    for temp_type in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
        y = df_clean[temp_type].values
        x = df_clean["Year_Num"].values

        if len(y) == 0:
            print(f"No valid {temp_type} data at station {station}. Skipping.")
            continue

        # Linear Regression
        model = LinearRegression().fit(x.reshape(-1, 1), y)
        linear_pred = model.predict(x.reshape(-1, 1))
        linear_slope = model.coef_[0]

        # Theil-Sen Estimator
        theil_slope, intercept, _, _ = theilslopes(y, x, 0.95)
        theil_pred = intercept + theil_slope * x

        trend_slopes[temp_type] = {
            "linear_slope": linear_slope,
            "theil_slope": theil_slope
        }

        # Plot original data
        plt.scatter(df_clean["Date"], y, color=colors[temp_type], alpha=0.4, s=10, label=f"{temp_type} Data")

        # # Plot Linear Regression Line
        # plt.plot(df_clean["Date"], linear_pred, color=trend_colors[temp_type], linewidth=1.5, linestyle="--", label=f"{temp_type} Linear")

        # Plot Theil–Sen Line
        plt.plot(df_clean["Date"], theil_pred, color=trend_colors[temp_type], linewidth=2, label=f"{temp_type} Theil–Sen")

        # Annotate Theil–Sen slope
        plt.text(df_clean["Date"].iloc[-1], theil_pred[-1], f"{theil_slope:.3f} °C/yr", 
                 fontsize=9, color=trend_colors[temp_type], fontweight='bold')

    plt.xlabel("Year")
    plt.ylabel("Temperature (°C)")
    plt.title(f"Temperature Trends at Station: {station_name} - using Theil–Sen Estimator")
    plt.legend()
    plt.grid(True)

    os.makedirs(output_folder_path, exist_ok=True)
    plot_path = os.path.join(output_folder_path, f"temperature_trend_{station}_using_theilsen.jpg")
    plt.savefig(plot_path, dpi=300)
    print(f"Theil–Sen trend plot saved at: {plot_path}")

    # plt.show()
    plt.close()

    return trend_slopes


def analyze_precipitation_trend_theilsen(output_folder_path, station, station_name):
    """
    Analyzes and plots daily precipitation trends over time using the Theil–Sen estimator.

    Parameters
    ----------
    output_folder_path : str
        Path where the processed precipitation dataset is stored.
    station : str
        Unique station identifier.
    station_name : str
        Readable station name for the plot.

    Returns
    -------
    theil_slope : float
        Estimated slope of the daily precipitation trend in mm/year.
    """
    file_path = os.path.join(output_folder_path, f"Daily_temperature_and_precipitation_station_{station}.xlsx")

    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return None

    df = pd.read_excel(file_path, sheet_name="Data")

    # Filter out rows without precipitation values
    if "Precipitation" not in df.columns:
        print(f"'Precipitation' column not found in the file.")
        return None

    df_clean = df.dropna(subset=["Precipitation"]).copy()

    # Combine Year, Month, Day into a single datetime column
    df_clean["Date"] = pd.to_datetime(df_clean[["Year", "Month", "Day"]])
    df_clean["Year_Num"] = df_clean["Year"].astype(int)

    if df_clean.empty:
        print(f"No valid precipitation data at station {station}.")
        return None

    # Predictor and response variables
    x = df_clean["Year_Num"].values
    y = df_clean["Precipitation"].values

    # Theil–Sen trend estimation
    theil_slope, intercept, _, _ = theilslopes(y, x, 0.95)
    theil_pred = intercept + theil_slope * x

    # Plotting
    plt.figure(figsize=(12, 6))
    plt.scatter(df_clean["Date"], y, alpha=0.3, s=10, color="dodgerblue", label="Daily Precipitation")
    plt.plot(df_clean["Date"], theil_pred, color="navy", linewidth=2, label="Theil–Sen Trend")

    plt.title(f"Precipitation Trends at Station: {station_name} - using Theil–Sen Estimator")
    plt.xlabel("Year")
    plt.ylabel("Precipitation (mm)")
    plt.legend()
    plt.grid(True)

    # Annotate slope (converted to mm/year)
    last_date = df_clean["Date"].iloc[-1]
    last_predicted = theil_pred[-1]
    plt.text(last_date, last_predicted, f"{theil_slope:.3f} mm/yr",
             fontsize=10, color="navy", fontweight="bold")

    # Save plot
    os.makedirs(output_folder_path, exist_ok=True)
    plot_path = os.path.join(output_folder_path, f"precipitation_trend_theilsen_{station}.jpg")
    plt.savefig(plot_path, dpi=300)
    print(f"Precipitation trend plot saved at: {plot_path}")

    # plt.show()
    plt.close()

    return theil_slope



def analyze_monthly_precipitation_trend_theilsen(output_folder_path, station, station_name):
    """
    Analyzes and plots monthly average precipitation trends over time using the Theil–Sen estimator.

    Parameters
    ----------
    output_folder_path : str
        Path where the processed precipitation dataset is stored.
    station : str
        Unique station identifier.
    station_name : str
        Readable station name for the plot.

    Returns
    -------
    theil_slope : float
        Estimated slope of the monthly precipitation trend in mm/year.
    """
    file_path = os.path.join(output_folder_path, f"Daily_temperature_and_precipitation_station_{station}.xlsx")

    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return None

    df = pd.read_excel(file_path, sheet_name="Data")

    if "Precipitation" not in df.columns:
        print(f"'Precipitation' column not found in the file.")
        return None

    df_clean = df.dropna(subset=["Precipitation"]).copy()
    df_clean["Date"] = pd.to_datetime(df_clean[["Year", "Month", "Day"]])
    
    if df_clean.empty:
        print(f"No valid precipitation data at station {station}.")
        return None

    # Group by Year and Month, then take the average
    df_clean["YearMonth"] = df_clean["Date"].dt.to_period("M")
    monthly_avg = df_clean.groupby("YearMonth")["Precipitation"].mean().reset_index()
    monthly_avg["Date"] = monthly_avg["YearMonth"].dt.to_timestamp()
    monthly_avg["Year_Fraction"] = monthly_avg["Date"].dt.year + (monthly_avg["Date"].dt.month - 1) / 12.0

    # Predictor and response variables
    x = monthly_avg["Year_Fraction"].values
    y = monthly_avg["Precipitation"].values

    # Theil–Sen estimator
    theil_slope, intercept, _, _ = theilslopes(y, x, 0.95)
    theil_pred = intercept + theil_slope * x

    # Plot
    plt.figure(figsize=(12, 6))
    plt.scatter(monthly_avg["Date"], y, alpha=0.4, s=15, color="skyblue", label="Monthly Avg Precipitation")
    plt.plot(monthly_avg["Date"], theil_pred, color="navy", linewidth=2, label="Theil–Sen Trend")

    plt.title(f"Precipitation Trends at Station: {station_name} - using Theil–Sen Estimator")
    plt.xlabel("Year")
    plt.ylabel("Monthly Avg Precipitation (mm)")
    plt.legend()
    plt.grid(True)

    # Annotate slope
    last_date = monthly_avg["Date"].iloc[-1]
    last_predicted = theil_pred[-1]
    plt.text(last_date, last_predicted, f"{theil_slope:.3f} mm/yr",
             fontsize=10, color="navy", fontweight="bold")

    # Save plot
    os.makedirs(output_folder_path, exist_ok=True)
    plot_path = os.path.join(output_folder_path, f"monthly_precipitation_trend_theilsen_{station}.jpg")
    plt.savefig(plot_path, dpi=300)
    print(f" Monthly precipitation trend plot saved at: {plot_path}")

    plt.close()

    return theil_slope


def convert_to_sef_with_metadata(
    df, station_info, temp_column, temp_type, source="Institut National pour l'Etude et la Recherche Agronomiques",
    link="", stat="point", units="C", observer="", hour=0
):
    """
    Convert a DataFrame containing temperature data into the Station Exchange Format (SEF, .tsv) using station metadata.

    Parameters
    ----------
    df : pandas.DataFrame
        DataFrame with 'Year', 'Month', 'Day', and the temperature column (`temp_column`).
    station_info : dict or Series
        Station metadata with keys: ID, name, latitude, longitude, altitude.
    temp_column : str
        Column name in `df` containing the temperature data.
    temp_type : str
        Type of temperature variable for SEF 'Vbl' field (e.g., "Tx", "Td").
    hour : int
        Hour of observation (e.g., 6 for 6 AM). Defaults to 0.
    source : str, optional
        Source of the data.
    link : str, optional
        URL link to data source.
    stat : str, optional
        Statistical type (e.g., 'point').
    units : str, optional
        Units of measurement, e.g., 'C' or 'mm'.
    observer : str, optional
        Observer or software metadata.

    Returns
    -------
    tuple
        - sef_headers: dict
        - sef_df: pandas.DataFrame (SEF-compliant data rows)
    """

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

    sef_df = pd.DataFrame({
        "Year": df["Year"],
        "Month": df["Month"],
        "Day": df["Day"],
        "Hour": [hour] * len(df),
        "Minute": [0] * len(df),
        "Period": [0] * len(df),
        "Value": df[temp_column].fillna("NaN"),
        "Meta": [""] * len(df)
    })

    sef_column_order = ["Year", "Month", "Day", "Hour", "Minute", "Period", "Value", "Meta"]
    sef_df = sef_df[sef_column_order]

    return sef_headers, sef_df


def data_formatting(input_folder_path, output_folder_path, metadata_file_path, station, date_column,
                    header_rows, multi_day_totals, multi_day_averages, excluded_rows,
                    additional_excluded_rows, final_totals_rows, uncertainty_margin):

    # Define your manually assigned column indices (based on Excel structure)
    max_temp_column_idx = 4   # 'D'
    min_temp_column_idx = 5   # 'E'
    avg_temp_column_idx = 6   # 'F'
    precip_column_idx   = 11  # 'K'
    dry06_column_idx    = 12  # 'L'
    wet06_column_idx    = 13  # 'M'
    dry15_column_idx    = 17  # 'Q'
    wet15_column_idx    = 18  # 'R'
    dry18_column_idx    = 22  # 'V'
    wet18_column_idx    = 23  # 'W'

    column_metadata = {
    "Max_Temperature": {"col": max_temp_column_idx, "vbl": "Tx", "hour": 0, "units": "C", "stat": "maximum"},
    "Min_Temperature": {"col": min_temp_column_idx, "vbl": "Tn", "hour": 0, "units": "C", "stat": "minimum"},
    "Avg_Temperature": {"col": avg_temp_column_idx, "vbl": "Ta", "hour": 0, "units": "C", "stat": "mean"},
    "Precipitation":   {"col": precip_column_idx,    "vbl": "rr", "hour": 6, "units": "mm", "stat": "sum"},
    "Dry_bulb_temp_06h00":  {"col": dry06_column_idx,     "vbl": "ta", "hour": 6, "units": "C", "stat": "point"},
    "Dry_bulb_temp_15h00":  {"col": dry15_column_idx,     "vbl": "ta", "hour": 15, "units": "C", "stat": "point"},
    "Dry_bulb_temp_18h00":  {"col": dry18_column_idx,     "vbl": "ta", "hour": 18, "units": "C", "stat": "point"},
    "Wet_bulb_temp_06h00":  {"col": wet06_column_idx,     "vbl": "tb", "hour": 6, "units": "C", "stat": "point"},
    "Wet_bulb_temp_15h00":  {"col": wet15_column_idx,     "vbl": "tb", "hour": 15, "units": "C", "stat": "point"},
    "Wet_bulb_temp_18h00":  {"col": wet18_column_idx,     "vbl": "tb", "hour": 18, "units": "C", "stat": "point"},
        }   ## Check https://datarescue.climate.copernicus.eu/variablenames for other variable names


    # Load station metadata
    station_metadata = load_station_metadata(metadata_file_path)
    station = str(station)
    station_metadata['ID'] = station_metadata['ID'].astype(str).str.strip().str.zfill(3)

    # Filter and extract station info
    station_info_df = station_metadata[station_metadata['ID'] == station]
    if station_info_df.empty:
        raise ValueError(f"No metadata found for station ID {station}")

    station_info = station_info_df.iloc[0]  # This is a Series, used in SEF headers
    station_name = station_info['name']  # Used for plot titles

    # Adjust excluded rows
    if multi_day_totals and not multi_day_averages:
        pass
    elif multi_day_totals and multi_day_averages:
        excluded_rows += additional_excluded_rows
    elif not multi_day_totals:
        excluded_rows = final_totals_rows

    # Convert column letter (e.g. 'B') to column index (1-based)
    date_column_idx = ord(date_column.upper()) - ord('A') + 1

    # Dictionary to store data for each variable
    data_by_variable = {var: [] for var in column_metadata}

    # Loop through Excel files
    for filename in os.listdir(input_folder_path):
        if filename.endswith(".xlsx"):
            year, month = extract_date_from_filename(filename)
            if year and month:
                file_path = os.path.join(input_folder_path, filename)
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active

                for row_num in range(header_rows + 1, ws.max_row + 1):
                    if row_num in excluded_rows:
                        continue

                    day_cell = ws.cell(row=row_num, column=date_column_idx)
                    if not day_cell.value:
                        continue
                    day = int(day_cell.value)

                    for var_name, meta in column_metadata.items():
                        col = meta["col"]
                        cell = ws.cell(row=row_num, column=col)
                        val = cell.value if is_highlighted_green(cell, 'FF6DCD57') else 'NaN'
                        data_by_variable[var_name].append([year, month, day, val])

    # Create merged DataFrame from collected data
    merged_df = pd.DataFrame()
    for key in data_by_variable:
        df_key = pd.DataFrame(data_by_variable[key], columns=["Year", "Month", "Day", key])
        if merged_df.empty:
            merged_df = df_key
        else:
            merged_df = pd.merge(merged_df, df_key, on=["Year", "Month", "Day"], how="outer")

    # Create full date range to ensure all days are present
    years_months = merged_df[["Year", "Month"]].drop_duplicates()
    full_dates = []
    for _, row in years_months.iterrows():
        y, m = row["Year"], row["Month"]
        for d in range(1, pd.Period(f'{y}-{m}').days_in_month + 1):
            full_dates.append([y, m, d])
    complete_df = pd.DataFrame(full_dates, columns=["Year", "Month", "Day"])
    merged_df = pd.merge(complete_df, merged_df, on=["Year", "Month", "Day"], how="left")
    merged_df = merged_df.sort_values(["Year", "Month", "Day"])

    # Convert values to numeric
    for col in merged_df.columns:
        if col not in ["Year", "Month", "Day"]:
            merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')

    # Detect outliers for temperature variables only (not precipitation or wet bulb)
    temp_vars = ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]
    for column in temp_vars:
        if column not in merged_df.columns:
            continue

        std = merged_df[column].std()
        mean = merged_df[column].mean()
        merged_df[f"{column}_Flag"] = ""

        # Condition 1: value > 3 standard deviations
        for i in range(len(merged_df)):
            if abs(merged_df.loc[i, column] - mean) > 3 * std:
                merged_df.loc[i, column] = np.nan
                merged_df.loc[i, f"{column}_Flag"] = "Condition 1" # Flag as Condition 1 (current value as an outlier)

        # Condition 2: sharp spike flip-flop
        merged_df['std_diff'] = (merged_df[column] - mean) / std
        for i in range(1, len(merged_df) - 1):
            prev_sd = merged_df.loc[i - 1, 'std_diff'] if not pd.isna(merged_df.loc[i - 1, 'std_diff']) else 0
            curr_sd = merged_df.loc[i, 'std_diff']
            next_sd = merged_df.loc[i + 1, 'std_diff'] if not pd.isna(merged_df.loc[i + 1, 'std_diff']) else 0
            if not pd.isna(curr_sd) and (
                (prev_sd < -4 and curr_sd > 4 and next_sd < -4) or
                (prev_sd > 4 and curr_sd < -4 and next_sd > 4)
            ):
                merged_df.loc[i, column] = np.nan
                merged_df.loc[i, f"{column}_Flag"] = "Condition 2" # Flag as Condition 2 (current value as an outlier)
        merged_df.drop(columns=['std_diff'], inplace=True)

    # Update data_by_variable with cleaned values
    for var in data_by_variable:
        if var in merged_df.columns:
            records = merged_df[["Year", "Month", "Day", var]].values.tolist()
            cleaned = [[int(y), int(m), int(d), v if not pd.isna(v) else "NaN"]
                    for y, m, d, v in records]
            data_by_variable[var] = cleaned


    # Export to SEF per variable
    for var_name, records in data_by_variable.items():
        if not records:
            continue

        df = pd.DataFrame(records, columns=["Year", "Month", "Day", "Value"])
        meta = column_metadata[var_name]

        sef_headers, sef_df = convert_to_sef_with_metadata(
        df=df,
        station_info=station_info,
        temp_column="Value",
        temp_type=meta["vbl"],
        hour=meta["hour"],
        units=meta["units"],
        stat=meta["stat"])

        sef_file = os.path.join(output_folder_path, f"SEF_station_{station}_{var_name}.tsv")
        with open(sef_file, 'w') as f:
            for hkey, hval in sef_headers.items():
                f.write(f"{hkey}\t{hval}\n")
            sef_df.to_csv(f, sep='\t', index=False) # Write the main SEF data with tab separation and include the header row for data columns


    # Save all cleaned variables as Excel files
    for var_name in data_by_variable:
        df_var = pd.DataFrame(data_by_variable[var_name], columns=["Year", "Month", "Day", var_name])
        df_var.to_excel(os.path.join(output_folder_path, f"{var_name}_timeseries.xlsx"), index=False)

    # Add Date column
    merged_df['Date'] = pd.to_datetime(merged_df[['Year', 'Month', 'Day']])

    #***PLOTTING***
    # === Plot 1: Combined Temp + Precip ===
    plot_df = merged_df.set_index('Date')[['Max_Temperature', 'Min_Temperature', 'Avg_Temperature', 'Precipitation']]
    fig, ax = plt.subplots(figsize=(10, 6))
    plot_df['Max_Temperature'].plot(ax=ax, label='Max Temp', color='red')
    ax.fill_between(plot_df.index, plot_df['Max_Temperature'] - uncertainty_margin, plot_df['Max_Temperature'] + uncertainty_margin, color='red', alpha=0.2)
    plot_df['Min_Temperature'].plot(ax=ax, label='Min Temp', color='blue')
    ax.fill_between(plot_df.index, plot_df['Min_Temperature'] - uncertainty_margin, plot_df['Min_Temperature'] + uncertainty_margin, color='blue', alpha=0.2)
    plot_df['Avg_Temperature'].plot(ax=ax, label='Avg Temp', color='orange')
    ax.fill_between(plot_df.index, plot_df['Avg_Temperature'] - uncertainty_margin, plot_df['Avg_Temperature'] + uncertainty_margin, color='orange', alpha=0.2)
    ax_precip = ax.twinx()
    ax_precip.scatter(plot_df.index, plot_df['Precipitation'], marker='x', color='cornflowerblue', label='Precipitation', alpha=0.7)
    ax.set_title(f"Temperature & Precipitation at Station: {station_name}")
    ax.set_ylabel("Temperature (°C)")
    ax_precip.set_ylabel("Precipitation (mm)")
    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax_precip.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
    ax.grid(True)
    plt.tight_layout()
    plt.savefig(os.path.join(output_folder_path, f"combined_temperature_and_precipitation_plot.jpg"))
    # plt.show()
    plt.close()

    # === Plot 2: Only Temperature ===
    fig, ax = plt.subplots(figsize=(10, 6))
    for col, color in zip(['Max_Temperature', 'Min_Temperature', 'Avg_Temperature'], ['red', 'blue', 'orange']):
        ax.plot(merged_df['Date'], merged_df[col], label=col.replace("_", " "), color=color)
        ax.fill_between(merged_df['Date'], merged_df[col] - uncertainty_margin, merged_df[col] + uncertainty_margin, alpha=0.2, color=color)
    ax.set_title(f"Temperature Time Series at Station: {station_name}")
    ax.set_xlabel("Date")
    ax.set_ylabel("Temperature (°C)")
    ax.legend()
    ax.grid(True)
    plt.tight_layout()
    plt.savefig(os.path.join(output_folder_path, f"temperature_plot_station_{station}.jpg"))
    # plt.show()
    plt.close()

    # === Plot 3: Precipitation ===
    # Filter out 0 mm precipitation
    non_zero_precip = merged_df[merged_df['Precipitation'] > 0]
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.scatter(non_zero_precip['Date'], non_zero_precip['Precipitation'], color='cornflowerblue', marker='x', alpha=0.7)
    ax.set_title(f"Precipitation at Station: {station_name}")
    ax.set_xlabel("Date")
    ax.set_ylabel("Precipitation (mm)")
    ax.grid(True)
    plt.tight_layout()
    plt.savefig(os.path.join(output_folder_path, f"precipitation_plot_station_{station}.jpg"))
    # plt.show()
    plt.close()

    # === Plot 4: Yearly Temperature Trends with 5-day Rolling Average ===
    for column in ["Max_Temperature", "Min_Temperature", "Avg_Temperature"]:
        cmap = {'Max_Temperature': 'Reds', 'Min_Temperature': 'Blues', 'Avg_Temperature': 'Oranges'}[column]
        num_years = merged_df['Year'].nunique()
        colors = [plt.cm.get_cmap(cmap, num_years)(i) for i in range(num_years)]
        year_to_color = dict(zip(sorted(merged_df['Year'].unique()), colors))

        fig, ax = plt.subplots(figsize=(12, 6))
        for year in sorted(merged_df['Year'].unique()):
            df_year = merged_df[merged_df['Year'] == year].copy()
            df_year[f'{column}_5day'] = df_year[column].rolling(window=5, min_periods=1).mean()
            ax.plot(df_year['Date'].dt.dayofyear, df_year[f'{column}_5day'], label=str(year), color=year_to_color[year])
            ax.text(df_year['Date'].dt.dayofyear.iloc[-1], df_year[f'{column}_5day'].iloc[-1], str(year),
                    fontsize=8, ha='left', color=year_to_color[year])
        ax.set_xticks([1, 32, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335])
        ax.set_xticklabels(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
        ax.set_title(f"{column.replace('_', ' ')} (5-day rolling) Trends at Station: {station_name}")
        ax.set_xlabel("Month")
        ax.set_ylabel(f"{column.replace('_', ' ')} (°C)")
        ax.legend(loc="upper left", fontsize=8, ncol=3)
        ax.grid(True)
        plt.tight_layout()
        plt.savefig(os.path.join(output_folder_path, f"{column}_trends_5day_station_{station}.jpg"))
        # plt.show()
        plt.close()

    # === Plot 5: Dry and Wet Bulb Plots ===
    plot_sets = {
        "Dry bulb temperatures": ["Dry_bulb_temp_06h00", "Dry_bulb_temp_15h00", "Dry_bulb_temp_18h00"],
        "Wet bulb temperatures": ["Wet_bulb_temp_06h00", "Wet_bulb_temp_15h00", "Wet_bulb_temp_18h00"],
    }

    for title, columns in plot_sets.items():
        fig, ax = plt.subplots(figsize=(10, 6))
        for col in columns:
            if col in merged_df.columns:
                merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')
                ax.plot(merged_df['Date'], merged_df[col], label=col.replace('_', ' '))
        ax.set_title(f"{title} at Station: {station_name}")
        ax.set_xlabel("Date")
        ax.set_ylabel("Temperature (°C)")
        ax.legend()
        ax.grid(True)
        plt.tight_layout()
        plt.savefig(os.path.join(output_folder_path, f"{title.replace(' ', '_').lower()}_plot_station_{station}.jpg"))
        # plt.show()
        plt.close()
    


    # === Plot 6: Trend lines with slope annotation ===
    # Save the cleaned full dataset (needed for trend analysis)
    trend_excel_path = os.path.join(output_folder_path, f"Daily_temperature_and_precipitation_station_{station}.xlsx")
    merged_df.to_excel(trend_excel_path, index=False, sheet_name="Data")
    
    # Analyze and plot temperature trends
    # 6.1: Linear Regression
    temperature_trend_slopes_with_linear_regression = analyze_temperature_trends_with_linear_regression(output_folder_path, station, station_name)
    # 6.2: Theil-Sen Estimator
    temperature_trend_slopes_with_theilsen = analyze_temperature_trends_with_theilsen(output_folder_path, station, station_name)
    # Precipitation_trends
    precipitation_trend_slopes_with_theilsen = analyze_monthly_precipitation_trend_theilsen(output_folder_path, station, station_name)

