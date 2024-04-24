import os
import openpyxl
from openpyxl.styles import PatternFill
import shutil
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment

# Setting up the current working directory; for both the input and output folders
cwd = os.getcwd()

def remove_single_quotes_from_numeric_cells(file_path):
    # Open het Excel-bestand
    wb = openpyxl.load_workbook(file_path)
    
    # Ga naar elk blad in het werkboek
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        # Itereer over alle cellen in het blad
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                # Controleer of de cel een numerieke waarde bevat en een string is
                if isinstance(cell.value, str):
                    # Verwijder enkele aanhalingstekens uit de celwaarde
                    cell.value = cell.value.replace("'", "")

    # Sla het bijgewerkte Excel-bestand op
    wb.save('table_1311_without_hyphens.xlsx')


def is_string_convertible_to_float(value):
    if value is None: # Check to handle None cases (empty cells)
        return False
    try:
        float(value)
        return True
    except ValueError:
        return False

def count_decimal_points(string): # Function to count the number of decimal points in a string
    count = 0
    for char in string:
        if char == '.':
            count += 1
    return count

    
def check_if_more_than_one_decimal(string):
    return string.count('.') == 2

def highlight_change(color, worksheet_and_cell_coordinate, filename):
    '''
    color: string
    worksheet_and_cell_coordinate: variable in the form worksheet[cell_coordinate] 
    filename: variable
    '''
    
    # Highlight cells with strings instead of floats
    highlighting_color = color # Highlighting color of choice
    highlighting_strings = PatternFill(start_color = highlighting_color, end_color = highlighting_color, fill_type = 'solid')
    cell_to_highlight = worksheet_and_cell_coordinate
    cell_to_highlight.fill = highlighting_strings
    # # save Excel file
    # workbook.save(filename)
    # return filename

    
def post_processing(pre_processed_excel_file):
    # Open the original Excel file
    workbook = openpyxl.load_workbook(pre_processed_excel_file)
    worksheet = workbook.active

    # Path to save a new copy of the workbook for post-processing
    new_version_of_file = 'quality_controlled_data_table_copy.xlsx'

    # Save the original workbook to ensure it's on disk
    original_path = 'original_transcribed_data.xlsx'
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