import os
import openpyxl
from openpyxl.styles import PatternFill
import shutil
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment
import pandas as pd

# Setting up the current working directory; for both the input and output folders
cwd = os.getcwd()



def is_string_convertible_to_float(value):
    '''
    # Check if the value in the cell can be converted to a float. This will ensure that calculations in the pprocessing of this data will be done.

    Parameters
    --------------
    value: value within a cell in an Ms Excel sheet

    Returns
    -------------- 
    value: convertible value within a cell. Passes the check.

    '''
    if value is None: # Check to handle None cases (empty cells)
        return False
    try:
        float(value)
        return True
    except ValueError:
        return False

def count_decimal_points(string): # Function to count the number of decimal points in a string
    '''
    # Counts the number of decimal points in a string in a cell. This is a check to avoid transcribed numbers with more than one decimal point

    Parameters
    --------------
    string: value within a cell in an Ms Excel sheet

    Returns
    -------------- 
    count: number of decimal points in a string in a cell
    '''

    count = 0
    for char in string:
        if char == '.':
            count += 1
    return count


def highlight_change(color, worksheet_and_cell_coordinate, filename):
    '''Highlights a cell in an Excel worksheet to show whether: (1) a change has been made in the processing of the transcribed data to correct an error, (2) certain values have been confirmed as correctly transcribed, or (3) certain values have been confirmed as wrongly transcribed

    Parameters
    --------------
    color: String. Selected color depending on the check. See table: Key_for_post_processed_data_sheets in the docs folders
    worksheet_and_cell_coordinate: cell to highlight
    filename: Name/Location of excel file

    Returns
    -------------- 
    highlighted cells in the excel file

    '''
    
    # Highlight cells with strings instead of floats
    highlighting_color = color # Highlighting color of choice
    highlighting_strings = PatternFill(start_color = highlighting_color, end_color = highlighting_color, fill_type = 'solid')
    cell_to_highlight = worksheet_and_cell_coordinate
    cell_to_highlight.fill = highlighting_strings
    # # save Excel file
    # workbook.save(filename)
    # return filename


def merge_excel_files(file1, file2, output_file, start_row, end_row):
    '''Merges two excel files (as a check): the transcribed excel organised in rows using the top coordinates of the bounding boxes and that organised in rows using the mid point coordinates.

    Parameters
    --------------
    file1: Excel sheet. Preprocessed transcribed data organised in rows using the mid point coordinates of the bounding boxes (contours).
    file2: Excel sheet. Preprocessed transcribed data organised in rows using the top coordinates of the bounding boxes(contours).
    output_file: Path. Location to store the output excel sheet. Merged file of file1 and file2 above to cross check to ensure propoer placement of cells in their rescpective rows
    start_row: Integer. Start row (beneath the headers)
    end_row: Integer. Last row

    Returns
    -------------- 
    Merged excel file: Now the pre-processed excel file.

    '''

    # Load the Excel files into DataFrames, ensuring they include headers if present
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Load headers separately if you need to prepend them later
    headers = pd.read_excel(file1, header= 0, nrows=3)  # Read only the first three rows for headers


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


    
def post_processing(pre_processed_excel_file, postprocessed_data_dir, month_filename):
    '''Post processing of the transcribed data. Here we make checks such as outlier detection (e.g. using thresholds for temperatures, variance, etc,) and ad-hoc corrections

    Parameters
    --------------
    pre_processed_excel_file: Excel sheet. Preprocessed transcribed data.
    postprocessed_data_dir: Path. Location to store the final postprocessed excel file.
    month_filename: String. Naming of output files with station metadata as in original images of climate data sheets

    Returns
    -------------- 
    new_workbook: Postprocessed excel file of transcribed climate data.

    '''
    # Open the original Excel file
    workbook = openpyxl.load_workbook(pre_processed_excel_file)
    worksheet = workbook.active

    # Path to save a new copy of the workbook for post-processing
    new_version_of_file = f'{postprocessed_data_dir}\{month_filename}_post_processed.xlsx'

    # Save the original workbook to ensure it's on disk
    original_path = f'src\output\original_transcribed_data.xlsx'
    workbook.save(original_path)

    # Copy the original file to a new version for post-processing
    shutil.copy2(original_path, new_version_of_file)

    new_workbook = openpyxl.load_workbook(new_version_of_file)
    new_worksheet = new_workbook.active # To select the first worksheet of the workbook without requiring its name

    workbook.close() # Close original transcribed workbook

    

    # Inorder to already avoud very large values right from the start, here we edit values that have more than 4 digits (thousands) in certain rows where we know it is impossible
    # List of rows to exclude from processing
    excluded_rows = [1, 2, 8, 15, 22, 29, 36, 44,46] # Headers and 5 day totals/ totals
    excluded_columns = [3, 15, 20, 25]  # Example of middle columns to exclude: with U
    # Iterate over all rows in the worksheet
    for row in new_worksheet.iter_rows(min_col=3, max_col=new_worksheet.max_column-1):
        # Get the row index from the first cell of the row
        if row[0].row not in excluded_rows:
            for cell in row:
                if cell.column not in excluded_columns:
                    if is_string_convertible_to_float(cell.value): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                        # Convert to string to check length and manipulate
                        str_value = str(cell.value)
                        if len(str_value) > 4: # Thousands. Here i chose 4 because it seems the strings are mad up up numbers and a newline \n character. Check this when you use 3 instead
                            # Remove the first digit and convert back to the original type
                            new_value = str_value[1:]  # Remove the first character
                            cell.value = type(cell.value)(new_value)  # Convert back to int or float
                            highlight_change('FF9933', cell, new_version_of_file) #FF9933 is 2 ## To highlight manipulation of transcribed data
                            # Save Excel file
                            new_workbook.save(new_version_of_file)
    
    # Save Excel file
    new_workbook.save(new_version_of_file)


    # Since values transcribed ignore decimal points for all cells yet all the cell values initially are in decimals, here we Iterate over all rows, starting from row 4 to skip header rows
    for row in new_worksheet.iter_rows(min_row=4, max_row=new_worksheet.max_row, min_col=3, max_col=new_worksheet.max_column-1): # For columns: avoid the first two columns and the last column (2nd with Date, and last also with the Date)
        for cell in row:
            if is_string_convertible_to_float(cell.value): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                cell.value = (float(cell.value))/10 # Divide the cell value by 10
                new_workbook.save(new_version_of_file) # Save the modified workbook

    # Label Date, Total and Average rows
    row_labels = ["1","2", "3", "4", "5", "Tot.", "Moy.", "6", "7", "8", "9", "10", "Tot.", "Moy.", "11", "12", "13", "14", "15", "Tot.", "Moy.", "16", "17", "18", "19", "20", "Tot.", "Moy.", "21", "22", "23", "24", "25", "Tot.", "Moy.", "26", "27", "28", "29", "30", "31", "Tot.", "Moy.", "Tot.", "Moy."]
    # Update the cells in the second and last column with the new values
    columns = [2, 27]
    for col in columns:
        for i, value in enumerate(row_labels, start=3):
            cell = new_worksheet.cell(row=i, column=col)
            cell.value = value
            new_workbook.save(new_version_of_file) # Save the modified workbook
    # Save Excel file
    new_workbook.save(new_version_of_file)

    # Delete all text in the first column (No de la pentade)
    # Calculate the number of rows in the worksheet
    max_row = new_worksheet.max_row
    # Iterate from row 4 to the last row (Rows 1-3 are headers)
    for row in range(3, max_row + 1):
        cell = new_worksheet.cell(row=row, column=1)
        cell.value = None  # Clear the content of the cell
        # Save Excel file
    new_workbook.save(new_version_of_file)
    

    # for row in range(4,50):
    #     for column in range:

    #         cell_data = new_worksheet[]


    # Creating thresholds for temperature values
    Maximum_Temperature_Threshold = 35  # Max reported temperatures during 1950-1959 were 30-31 degC + increasing temperatures in 1960-1990 approximated at 0.60°C to 1.62°C per 30 yr period (Alsdorf et.al, 2016)
    Minimum_Temperature_Threshold = 10  # Min reported temperatures during 1950-1959 were 18-20 degC  et.al, 2016)       

    # Cell coordinates to check the totals
    # Rows
    rows = [8, 15, 22, 29, 36, 44]
    #rows = [9] #just for trial

    for row in rows:

        # Columns
        columns = ['D', 'E', 'F', 'K'] # Where these represent [ Max Temperature, Min Temperature, Average Temperature, Precipitation] 
        #columns = ['E'] #just for trial

        for column in columns:
            # Get the value of the '5/6-day Total' that was transcribed. Note Months with 31 days ahve some 6 day todays considering 26th to 31st
            cell_for_5_day_total_retrieved = new_worksheet[column + str(row)]
            value_in_cell_for_5_day_total_retrieved = cell_for_5_day_total_retrieved.value
            value_in_cell_for_5_day_total_retrieved_as_string = str(value_in_cell_for_5_day_total_retrieved) # Convert it to a string. To ensure uniformity
            
            if ',' or '.' in value_in_cell_for_5_day_total_retrieved_as_string:
                value_in_cell_for_5_day_total_retrieved_as_string= value_in_cell_for_5_day_total_retrieved_as_string.replace(',','.') # Replace comma with decimal point.  Second check below
                # cell_for_5_day_total_retrieved.value = (float(value_in_cell_for_5_day_total_retrieved_as_string))
                # Checking the number of decimal points in the string
                number_of_decimal_points = count_decimal_points(value_in_cell_for_5_day_total_retrieved_as_string)
                # *********Highlight this change in orange to identify the possible error of a decimal point appearing as a comma
                # *********highlight_change('FFA500', cell_for_5_day_total_retrieved, filename) # FFA500 is Orange
                
                if number_of_decimal_points == 2: #If 2 decimal points, remove the first one.
                    new_value_in_cell_for_5_day_total_retrieved_as_string = value_in_cell_for_5_day_total_retrieved_as_string.replace('.','',1) # Eliminate the first decimal point as a first check                
                    if is_string_convertible_to_float(new_value_in_cell_for_5_day_total_retrieved_as_string): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                        cell_for_5_day_total_retrieved.value = (float(new_value_in_cell_for_5_day_total_retrieved_as_string))
                        #value_in_cell_for_5_day_total_retrieved = new_value_in_cell_for_5_day_total_retrieved_as_string
                        # Highlight this change in yellow to identify the possible error of a decimal point appearing twice
                        highlight_change('FFFF00', cell_for_5_day_total_retrieved, new_version_of_file) # FFFF00 is Yellow
                        # Save Excel file
                        new_workbook.save(new_version_of_file)
                    else:
                        # Highlight cell to identify this error of a string instead of a float
                        highlight_change('FF0000', cell_for_5_day_total_retrieved, new_version_of_file) #FF0000 is Red
                        print(column + str(row)+ ' has a word instead of a value in the original transcribed data. Check this')
                        # Save Excel file
                        new_workbook.save(new_version_of_file)
                
                if number_of_decimal_points > 2:
                    # Highlight this change orange to identify this error of more than 2 decimal points after removing the first one.
                    highlight_change('FF9933', cell_for_5_day_total_retrieved, new_version_of_file) #'FF9933' is Orange 
                    print(column + str(row)+ ' has more decimal points that one in the original transcribed data. Check this')
                    # Save Excel file
                    new_workbook.save(new_version_of_file)
                    
                else:
                    if is_string_convertible_to_float(value_in_cell_for_5_day_total_retrieved_as_string): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                        cell_for_5_day_total_retrieved.value = (float(value_in_cell_for_5_day_total_retrieved_as_string))
                        #value_in_cell_for_5_day_total_retrieved = value_in_cell_for_5_day_total_retrieved_as_string
                        new_workbook.save(new_version_of_file)
                    else:
                        # Highlight cell to identify this error of a string instead of a float
                        highlight_change('FF0000', cell_for_5_day_total_retrieved, new_version_of_file) #FF0000 is Red
                        print(column + str(row)+ ' has a word instead of a value in the original transcribed data. Check this')
                        # Save Excel file
                        new_workbook.save(new_version_of_file)
                        
            
            print(value_in_cell_for_5_day_total_retrieved_as_string)


            # Select the 5 cells just above the '5-day/6-day Total' in the transcribed file
            
            # Note, while most totals are based on 5-day values, some months with 31 days have 6-day totals or the last days. 
            # Thus here, the definition of '5-day total' is used for both the 5-day and 6-day totals for simplicity since the value is only used as Quality control check of the daily values
            
            if row in [8, 15, 22, 29, 36]: # rows with only 5-day totals for all months
                offset_cells = 5
            
            else: # last rows with 6-day totals during months with 31 days or EVEN less than 5-days totals during February
                offset_cells = 6
            
            blank_cells = [] # Blank cells (None Type) for dates on the sheet that dont correspond to calender dates for example cells for Feb (29th), 30th and 31st that don't exist
            
            cell_values_for_the_5_days = cell_for_5_day_total_retrieved.offset(row = -offset_cells, column = 0)
            list_of_cell_values_for_the_5_days = []
            for cells in range(offset_cells): 
                cell_values_for_5_days_retrieved = cell_values_for_the_5_days.offset(row = cells, column = 0).value
                
                cell_coordinate = column + str(row + cells - offset_cells) # Identify the particular cell coordinate
                cell_coordinate_in_worksheet = new_worksheet[cell_coordinate]
                
                if ',' or '.' in cell_values_for_5_days_retrieved:
                    value_in_cell_retrieved_as_string= str(cell_values_for_5_days_retrieved).replace(',','.') # Replace comma with decimal point. Second check below
                    # cell_for_5_day_total_retrieved.value = (float(value_in_cell_for_5_day_total_retrieved_as_string))
                    # Checking the number of decimal points in the string
                    number_of_decimal_points = count_decimal_points(value_in_cell_retrieved_as_string)
                    # *********Highlight this change in orange to identify the possible error of a decimal point appearing as a comma
                    # *********highlight_change('FFA500', cell_for_5_day_total_retrieved, filename) # FFA500 is Orange
                    
                    
                    
                    if number_of_decimal_points == 2: #If 2 decimal points, remove the first one.
                        new_value_in_cell_retrieved_as_string = value_in_cell_retrieved_as_string.replace('.','',1) # Eliminate the first decimal point as a first check                
                        if is_string_convertible_to_float(new_value_in_cell_retrieved_as_string): # Checking if the value transcribed in the cell is convertible to a float to avoid strings
                            list_of_cell_values_for_the_5_days.append(float(new_value_in_cell_retrieved_as_string))
                            
                            cell_coordinate_in_worksheet.value = float(new_value_in_cell_retrieved_as_string)
                            #value_in_cell_retrieved_as_string = new_value_in_cell_for_5_day_total_retrieved_as_string
                            
                            # Highlight this change in yellow to identify the possible error of a decimal point appearing twice
                            highlight_change('FFFF00', cell_coordinate_in_worksheet, new_version_of_file) # FFFF00 is Yellow
                            # Save Excel file
                            new_workbook.save(new_version_of_file)
                            
                        else:
                            list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.         

                            #cell_coordinate = column + str(row + cells - 5)
                            print('Check the value within '+ cell_coordinate) # Report the cell with a string instead of a float
                            # Highlight cells with strings instead of floats
                            highlight_change('FF0000', cell_coordinate_in_worksheet, new_version_of_file) # FF0000 is Red. Original transcription had a word instead of a number. Check this.
                            # save Excel file
                            new_workbook.save(new_version_of_file)

                    if number_of_decimal_points > 2:
                        list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.         
                        
                        # Highlight this change red to identify this error of more than 2 decimal points after removing the first one.
                        highlight_change('FF0000', cell_coordinate_in_worksheet, new_version_of_file) #FF0000 is Red 
                        new_workbook.save(new_version_of_file)
                        print(cell_coordinate + ' has more decimal points that one in the original transcribed data. Check this')
                        new_workbook.save(new_version_of_file)

                    else:
                        if is_string_convertible_to_float(value_in_cell_retrieved_as_string): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                            list_of_cell_values_for_the_5_days.append(float(value_in_cell_retrieved_as_string))
                            cell_coordinate_in_worksheet.value = (float(value_in_cell_retrieved_as_string))
                            # save Excel file
                            new_workbook.save(new_version_of_file)
                        
                        
                        else:
                            list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.  
                            
                            if value_in_cell_retrieved_as_string == "":
                                highlight_change('FFFFFF', cell_coordinate_in_worksheet, new_version_of_file) #No Fill
                                new_workbook.save(new_version_of_file)
                            
                            else:
                                # Highlight this change red to identify strings.
                                highlight_change('FF0000', cell_coordinate_in_worksheet, new_version_of_file) #FF0000 is Red 
                                #print(cell_coordinate + ' has a string instead of a float in the original transcribed data. Check this')
                                # save Excel file
                                new_workbook.save(new_version_of_file)


                else:
                    if is_string_convertible_to_float(cell_values_for_5_days_retrieved): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                        list_of_cell_values_for_the_5_days.append(float(cell_values_for_5_days_retrieved))
                        cell_coordinate_in_worksheet.value = (float(cell_values_for_5_days_retrieved))
                        new_workbook.save(new_version_of_file)
                        
                    else:
                        list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.
                        
                        if str(cell_values_for_5_days_retrieved) == "":
                                highlight_change('FFFFFF', cell_coordinate_in_worksheet, new_version_of_file) #No Fill
                                new_workbook.save(new_version_of_file)
                            
                        else:
                            
                            # Highlight this change red to identify this error of more than 2 decimal points after removing the first one.
                            highlight_change('FF0000', cell_coordinate_in_worksheet, new_version_of_file) #FF0000 is Red 
                            #print(cell_coordinate + ' has a string instead of a float in the original transcribed data. Check this')
                            new_workbook.save(new_version_of_file)

                        
                        #if cell_values_for_5_days_retrieved == '':
                            # Leave blank
                            #list_of_cell_values_for_the_5_days.append(0.0) # Convert it to 0.0 to avoid errors in the summation.
                            #new_workbook.save(new_version_of_file)
                        
                        #else:
                            #list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.  

                            # Highlight this change red to identify this error of more than 2 decimal points after removing the first one.
                            #highlight_change('FF0000', cell_coordinate_in_worksheet, new_version_of_file) #FF0000 is Red 
                            #print(cell_coordinate + ' has a string instead of a float in the original transcribed data. Check this')
                            #new_workbook.save(new_version_of_file)

                # Blank cells
                if cell_values_for_5_days_retrieved is None or cell_values_for_5_days_retrieved == '':
                    blank_cells.append(1.0) # Count the number of blank cells representing no data or dates that dont exist on the calender for example cells for Feb (29th), 30th and 31st
                    highlight_change('FFFFFF', cell_coordinate_in_worksheet, new_version_of_file) #No Fill

                
                if column in ['D', 'E', 'F']: # Maximimum, Minimum and Average Temperatures
                    
                    # Creating thresholds for temperature values
                    #Maximum_Temperature_Threshold = 35  # Max reported temperatures during 1950-1959 were 30-31 degC + increasing temperatures in 1960-1990 approximated at 0.60°C to 1.62°C per 30 yr period (Alsdorf et.al, 2016)
                    #Minimum_Temperature_Threshold = 10  # Min reported temperatures during 1950-1959 were 18-20 degC  et.al, 2016)       
                    
                    if cell_coordinate_in_worksheet.value is None or cell_coordinate_in_worksheet.value == '':
                        # Highlight to show that cell is empty
                        highlight_change('FFFFFF', cell_coordinate_in_worksheet, new_version_of_file) ##FFFFFF is white.
                        new_workbook.save(new_version_of_file)
                    else: 
                        if is_string_convertible_to_float(cell_coordinate_in_worksheet.value): 
                            if float(cell_coordinate_in_worksheet.value) > Maximum_Temperature_Threshold:
                                
                                if float(cell_coordinate_in_worksheet.value) == "":
                                    highlight_change('FFFFFF', cell_coordinate_in_worksheet, new_version_of_file) #No Fill
                                    new_workbook.save(new_version_of_file)
                                else:
                                    # Highlight to show that value is out of the expected bounds
                                    new_value = (cell_coordinate_in_worksheet.value)/10.0  #try dividing by 10, to avoid very large values due to missing decimal point
                                    cell_coordinate_in_worksheet.value = new_value


                                    highlight_change('CC3300', cell_coordinate_in_worksheet, new_version_of_file) #CC3300 is Dark Red 


                                    new_workbook.save(new_version_of_file)
                
                
                    ### Doing the check again after manipulating the values
                    list_of_cell_values_for_the_5_days = []
                    for cells in range(offset_cells): 
                        cell_values_for_5_days_retrieved = cell_values_for_the_5_days.offset(row = cells, column = 0).value

                        cell_coordinate = column + str(row + cells - offset_cells) # Identify the particular cell coordinate
                        cell_coordinate_in_worksheet = new_worksheet[cell_coordinate]
                        
                        if cell_coordinate_in_worksheet.value is not None:
                            
                            if is_string_convertible_to_float(cell_values_for_5_days_retrieved): # Checking if the value transcribed in the cell is convertible to a float to avoid strings 
                                list_of_cell_values_for_the_5_days.append(float(cell_values_for_5_days_retrieved))
                                cell_coordinate_in_worksheet.value = (float(cell_values_for_5_days_retrieved))
                                new_workbook.save(new_version_of_file)

                            else:
                                list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.




                            #if cell_values_for_5_days_retrieved == '':
                                # Leave blank
                                #list_of_cell_values_for_the_5_days.append(0.0) # Convert it to 0.0 to avoid errors in the summation.
                                #new_workbook.save(new_version_of_file)

                            #else:
                                #list_of_cell_values_for_the_5_days.append(0.0) # Incase the value transcribed in the cell is a string, convert it to 0.0 to avoid errors in the summation.  

                                # Highlight this change red to identify this error of more than 2 decimal points after removing the first one.
                                #highlight_change('FF0000', cell_coordinate_in_worksheet, new_version_of_file) #FF0000 is Red 
                                #print(cell_coordinate + ' has a string instead of a float in the original transcribed data. Check this')
                                #new_workbook.save(new_version_of_file)
                    
                
                    
            
            total_from_cell_values_for_the_5_days = sum(list_of_cell_values_for_the_5_days)
            
            #Compare the Total of the 5-days/6-days transcribed and that calculated from the transcribed data
            if is_string_convertible_to_float(value_in_cell_for_5_day_total_retrieved_as_string):
                #if value_in_cell_for_5_day_total_retrieved ==  str(total_from_cell_values_for_the_5_days): 
                if format(float(value_in_cell_for_5_day_total_retrieved_as_string),'.1f') == format(float(total_from_cell_values_for_the_5_days), '.1f'):
                    highlight_change('6DCD57', new_worksheet[column + str(row)], new_version_of_file) #6DCD57 is Green. The total of the transcribed 5 days values is equal to the Total transcribed for the cells
                    new_workbook.save(new_version_of_file)
                    # highlight transcribed cells that lead to correct transcribed total
                    for i in range(row - offset_cells, row):
                        highlight_change('6DCD57', new_worksheet[column + str(i)], new_version_of_file) #6DCD57 is Green. The total of the transcribed 5 days values is equal to the Total transcribed for the cells
                        new_workbook.save(new_version_of_file)
                    print('The total of the transcribed 5 days values is equal to the Total transcribed for the cell; ' +str(column)+ str(row) + ' is OK')

                else:
                    highlight_change('75696F', new_worksheet[column + str(row)], new_version_of_file) #75696F is Grey. When transcribed 5-day total is not equal to total of the 5 days
                    new_workbook.save(new_version_of_file)
                    print('Check the Total transcribed at cell ' + str(column)+ str(row) +', or the transcribed 5 days values above, because the total of the transcribed 5 days values is not equal to the Total transcribed')
            
            
            # Compare the Mean of the 5-days/6-days transcribed and that calculated from the transcribed data
            cell_coordinate_in_worksheet_with_the_mean = new_worksheet[column + str(row+1)]
            mean_of_5_days_retrieved = cell_coordinate_in_worksheet_with_the_mean.value
            if mean_of_5_days_retrieved is None or mean_of_5_days_retrieved == '':  # Highlight empty cells
                # Highlight to show that cell is empty 
                highlight_change('FFC0CB', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #FFC0CB is Pink
                new_workbook.save(new_version_of_file)
            else:
                if is_string_convertible_to_float(mean_of_5_days_retrieved): 
                    
                    total_blank_cells = sum(blank_cells) # Count of the blank cells
                    
                    if format(float(mean_of_5_days_retrieved), '.1f') != format(total_from_cell_values_for_the_5_days/(offset_cells - total_blank_cells), '.1f'): # If the mean transcribed (retrieved from the transciption) is not equal to the calculated mean of the values , then highlight the cell
                        cell_coordinate_in_worksheet_with_the_mean.value = float(mean_of_5_days_retrieved)
                        highlight_change('75696F', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #75696F is Grey. When transcribed 5-day average is not equal to average of the 5 days
                        
                        if float(mean_of_5_days_retrieved) > Maximum_Temperature_Threshold:
                                # Highlight to show that value is out of the expected bounds
                                new_value = (float(mean_of_5_days_retrieved)/10.0)  #try dividing by 10, to avoid very large values due to missing decimal point
                                cell_coordinate_in_worksheet_with_the_mean.value = new_value
                                highlight_change('CC3300', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #CC3300 is Dark Red
                                
                                if format(float(new_value), '.1f') == format(total_from_cell_values_for_the_5_days/(offset_cells - total_blank_cells), '.1f'):
                                    highlight_change('6DCD57', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #6DCD57 is Green. When transcribed 5-day average is equal to average of the 5 days
                            
                        new_workbook.save(new_version_of_file)
                    else:
                        highlight_change('6DCD57', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #6DCD57 is Green. When transcribed 5-day average is equal to average of the 5 days
                        # highlight transcribed cells that lead to correct transcribed average
                        
                        for i in range(row - offset_cells, row):
                            highlight_change('6DCD57', new_worksheet[column + str(i)], new_version_of_file) #6DCD57 is Green. The average of the transcribed 5 days values is equal to the average transcribed for the cells
                            new_workbook.save(new_version_of_file)
                        
                        new_workbook.save(new_version_of_file)
                
            # DO it again! Compare the Mean of the 5-days/6-days transcribed and that calculated from the transcribed data
            cell_coordinate_in_worksheet_with_the_mean = new_worksheet[column + str(row+1)]
            mean_of_5_days_retrieved = cell_coordinate_in_worksheet_with_the_mean.value
            if mean_of_5_days_retrieved is None or mean_of_5_days_retrieved == '':  # Highlight empty cells
                # Highlight to show that cell is empty 
                highlight_change('FFC0CB', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #FFC0CB is Pink
                new_workbook.save(new_version_of_file)
            else:
                if is_string_convertible_to_float(mean_of_5_days_retrieved): 
                    
                    total_blank_cells = sum(blank_cells) # Count of the blank cells
                    
                    if format(float(mean_of_5_days_retrieved), '.1f') != format(total_from_cell_values_for_the_5_days/(offset_cells - total_blank_cells), '.1f'): # If the mean transcribed (retrieved from the transciption) is not equal to the calculated mean of the values , then highlight the cell
                        cell_coordinate_in_worksheet_with_the_mean.value = float(mean_of_5_days_retrieved)
                        highlight_change('75696F', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #75696F is Grey. When transcribed 5-day average is not equal to average of the 5 days
                        
                        if float(mean_of_5_days_retrieved) > Maximum_Temperature_Threshold:
                                # Highlight to show that value is out of the expected bounds
                                new_value = (float(mean_of_5_days_retrieved)/10.0)  #try dividing by 10, to avoid very large values due to missing decimal point
                                cell_coordinate_in_worksheet_with_the_mean.value = new_value
                                highlight_change('CC3300', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #CC3300 is Dark Red
                                
                                if format(float(new_value), '.1f') == format(total_from_cell_values_for_the_5_days/(offset_cells - total_blank_cells), '.1f'):
                                    highlight_change('6DCD57', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #6DCD57 is Green. When transcribed 5-day average is equal to average of the 5 days
                            
                        new_workbook.save(new_version_of_file)
                    else:
                        highlight_change('6DCD57', cell_coordinate_in_worksheet_with_the_mean, new_version_of_file) #6DCD57 is Green. When transcribed 5-day average is equal to average of the 5 days
                        # highlight transcribed cells that lead to correct transcribed average
                        
                        for i in range(row - offset_cells, row):
                            highlight_change('6DCD57', new_worksheet[column + str(i)], new_version_of_file) #6DCD57 is Green. The average of the transcribed 5 days values is equal to the average transcribed for the cells
                            new_workbook.save(new_version_of_file)
                        
                        new_workbook.save(new_version_of_file)


    

    # Insert a new row at the top for headers
    new_worksheet.insert_rows(1, amount = 1)
    # Define your headers (adjust as needed)
    headers = ["No de la pentade", "Date", "Bellani (gr. Cal/cm2) 6-6h", "Températures extrêmes", "", "", "", "", "Evaportation en cm3 6 - 6h", "", "Pluies en mm. 6-6h", "Température et Humidité de l'air à 6 heures", "", "", "", "", "Température et Humidité de l'air à 15 heures",  "", "", "", "", "Température et Humidité de l'air à 18 heures",  "", "", "", "", "Date"]
    # Add the headers to the first row
    for col_num, header in enumerate(headers, start=1):
        new_worksheet.cell(row=1, column=col_num, value=header)

        if header == "No de la pentade" or header == "Date" or header == "Bellani (gr. Cal/cm2) 6-6h" or header == "Pluies en mm. 6-6h":
            cell.alignment = Alignment(textRotation=90)
    
    # Merge cells for multi-column headers
    new_worksheet.merge_cells(start_row=1, start_column=1, end_row=3, end_column=1) #No de la pentade
    new_worksheet.merge_cells(start_row=1, start_column=2, end_row=3, end_column=2) #Date
    new_worksheet.merge_cells(start_row=1, start_column=3, end_row=3, end_column=3) #Bellani
    new_worksheet.merge_cells(start_row=1, start_column=4, end_row=1, end_column=8) #Températures extrêmes
    new_worksheet.merge_cells(start_row=1, start_column=9, end_row=1, end_column=10) #Evaportation
    new_worksheet.merge_cells(start_row=1, start_column=11, end_row=3, end_column=11) #Pluies
    new_worksheet.merge_cells(start_row=1, start_column=12, end_row=1, end_column=16) #Température et Humidité de l'air à 6 heures
    new_worksheet.merge_cells(start_row=1, start_column=17, end_row=1, end_column=21) #Température et Humidité de l'air à 15 heures
    new_worksheet.merge_cells(start_row=1, start_column=22, end_row=1, end_column=26) #Température et Humidité de l'air à 18 heures
    new_worksheet.merge_cells(start_row=1, start_column=27, end_row=3, end_column=27) #Date
    # subheaders
    new_worksheet.merge_cells(start_row=2, start_column=4, end_row=2, end_column=7) #Abri
    new_worksheet.merge_cells(start_row=2, start_column=9, end_row=2, end_column=10) #Piche
    new_worksheet.merge_cells(start_row=2, start_column=12, end_row=2, end_column=16) #(Psychromètre a aspiration)
    new_worksheet.merge_cells(start_row=2, start_column=17, end_row=2, end_column=21) #(Psychromètre a aspiration)
    new_worksheet.merge_cells(start_row=2, start_column=22, end_row=2, end_column=26) #(Psychromètre a aspiration)

    
    # Set up border styles for excel output
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'))

    # Loop through cells to apply borders
    for row in new_worksheet.iter_rows(min_row=1, max_row=new_worksheet.max_row, min_col=1, max_col=new_worksheet.max_column):
        for cell in row:
            cell.border = thin_border
    new_workbook.save(new_version_of_file)
    
    # Iterate through all cells and set the alignment
    for row in new_worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    new_workbook.save(new_version_of_file)

    new_workbook.close()

    return new_workbook