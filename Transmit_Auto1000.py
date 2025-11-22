import os
import re
import shutil
import glob
from openpyxl import load_workbook
import win32com.client

def find_excel_file():
    print("Searching for Transmittal Excel file...")        
    source_file = r"C:\Users\eddpa\Desktop\GoalFolder\Transmittal_TEMPLATE.xlsx" # Replace with template path used in company
    rootDirectory = os.path.abspath(".")
    rootKey = "."
    transmittal_excel = "Transmittal.xlsx"
    for currentDir, subDir, fileNames in os.walk(rootKey):
        for file in fileNames:
            transmittal_excel = os.path.join(currentDir, file)
            if file == "Transmittal.xlsx" and currentDir != rootKey:
                shutil.copy(transmittal_excel, rootDirectory)
                print("Transmittal Excel file found and moved into root directory.")                
                return os.path.join(rootDirectory, "Transmittal.xlsx")
            if file == "Transmittal.xlsx" and currentDir == rootKey:
                print("Transmittal Excel file found in root directory.")
                return os.path.join(rootDirectory, "Transmittal.xlsx")

    print("Transmittal Excel file not found. Copying from source...")
    try:
        shutil.copy(source_file, rootDirectory)
        print(f"Template Transmittal'{source_file}' copied successfully to '{rootDirectory}'")
    except FileNotFoundError:
        print("Error: The source file was not found.")
    except PermissionError:
        print("Error: You do not have permission to write to the destination.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    return os.path.join(rootDirectory, "Transmittal_TEMPLATE.xlsx")

def Catch_Drawings():
    print("Searching for drawings in current folder and subfolders...")    
    # Define pattern of file name that ends in .pdf and has a character or digit inside square brackets
    pattern = "*[[]?[]]*.pdf"
    
    # Store found PDF drawings in a list
    rawListDrawings = glob.glob(pattern)

    if rawListDrawings:
        print(f"Found {len(rawListDrawings)} drawings in the current folder.") 
        return rawListDrawings
    else:
        
        return None

def Extract_Names_From_Drawings():
    # Get raw list of PDF drawing names
    rawlist_PDF = Catch_Drawings()
    # Extract just drawing names from PDF list
    list_PDF = []
    #define pattern that catches any string content between "] " and ".pdf"
    pattern = r"\] (.*)\.pdf"
    if rawlist_PDF:
        for drawing in rawlist_PDF:
            match = re.search(pattern, drawing)
            if match:
                catched_name = match.group(1)
                list_PDF.append(catched_name)
        return list_PDF
    else:
        return None
    
def Request_Get_Date():
    rootDirectory = os.path.abspath(".")
    print("Choose a date for the transmittal (DD/MM/YY), from the following options: ")
    while True:
        # Print the menu options        
        print("1. Today's Date")
        print("2. Enter a Custom Date")

        # Prompt for user input
        choice = input("Enter your choice (1-3): ")
        # Activate choice
        if choice == '1':
            from datetime import datetime
            now = datetime.now()
            date_string = now.strftime("%d/%m/%y")
            break
        elif choice == '2':
            date_string = input("Enter the date in the following format DD/MM/YY : ")
            if len(date_string) != 10 or date_string[2] != '/' or date_string[5] != '/':
                print("Invalid date format. Please try again.")
                input("Press Enter to continue...")
                continue
            break
        elif choice == '3':
            print("Exiting the program. Goodbye!")
            break  # Exit the while loop
        else:
            print("Invalid choice. Please enter a number between 1 and 3.")
            input("Press Enter to continue...")
    # Load transmittal Excel file
    transmittal = find_excel_file()
    workbook = load_workbook(transmittal)

    # Check if 'CIVIL' sheet exists
    if 'CIVIL' in workbook.sheetnames:
        print("CIVIL sheet found.")
        worksheet = workbook['CIVIL']                       
    else:
        raise ValueError("Sheet 'CIVIL' not found in the Excel file.")

    # Get the date from user and overwrite empty date cells    
    dateStrings = date_string.split('/')
    dateParts = [int(s) for s in dateStrings]
    fileName_date = f"{dateParts[2]}{dateParts[1]}{dateParts[0]}"
    dateRow_index = 1    
    for date_cells in worksheet.iter_rows(min_row=dateRow_index, max_row=dateRow_index, min_col=5, max_col=30):    
        for cell in date_cells:
            if cell.value is None:
                print("New date cell is: " + cell.coordinate)
                for i, part in enumerate(dateParts):
                    target_cell = dateRow_index + i
                    worksheet.cell(row=target_cell, column=cell.column, value=part)                    
                break   

    # Save the modified workbook    
    output_filename = f"Transmittal {fileName_date}.xlsx"
    workbook.save(output_filename)
    print(f"Transmittal excel file saved as '{output_filename}'.")

    return os.path.join(rootDirectory, output_filename)

def Overwrite_Transmittal():
    # Load transmittal Excel file
    print("Loading new Transmittal Excel file...")
    transmittal = Request_Get_Date()
    workbook = load_workbook(transmittal)

    # Check if 'CIVIL' sheet exists. Select it if it does
    if 'CIVIL' in workbook.sheetnames:
        print("CIVIL sheet found.")
        worksheet = workbook['CIVIL']                       
    else:
        raise ValueError("Sheet 'CIVIL' not found in the Excel file.")      

    # Get revision column index from Excel by last date reference
    dateRow_index = 1
    for date_cells in worksheet.iter_rows(min_row=dateRow_index, max_row=dateRow_index, min_col=5, max_col=30):    
        for cell in date_cells:
            if cell.value is None:
                rev_column = cell.column - 1  # Revision column is the one before the empty cell
                break # If found, Exit inner loop
        if rev_column:
            break # Exit outer loop as well
    if not rev_column:
        print("Error: Could not determine the revision column.")
        return # Exit the function if NO revision column is found

    ######### Compare PDF vs Excel names, update revision when matched and add new drawings ##########
    # Get revision from raw list of PDF drawings
    rawlist_PDF = Catch_Drawings()

    if not rawlist_PDF:
        print("No drawings found to process.")
        return # Exit the function if no drawings found

    for file_Name in rawlist_PDF:
        # Define regex patterns to extract project number, revision, and drawing name
        pjtNo_pattern = r"(.*)-C" # Pattern to catch anything before "["
        dwg_pattern = r"\d{5}-(.*) \[" # Pattern to catch anything between 5 digits with a hyphen and " ["
        rev_pattern = r"\[(.*)\]" # Pattern to catch anything between "[ " and "]"        
        name_pattern = r"\] (.*)\.pdf" # Pattern to catch anything between "] " and ".pdf"
        pjtNo_match = re.search(pjtNo_pattern, file_Name)
        dwg_match = re.search(dwg_pattern, file_Name)
        rev_match = re.search(rev_pattern, file_Name)
        name_match = re.search(name_pattern, file_Name)
        
        # Skip file if pattern not found
        if not (pjtNo_match and dwg_match and rev_match and name_match):
            print(f"Could not parse all details from '{file_Name}'. Skipping.")
            continue # Skip to the next file
        
        # Strip found values from " " and store them
        project_No = pjtNo_match.group(1).strip() 
        drawing_No = dwg_match.group(1).strip() 
        revision = rev_match.group(1).strip() 
        drawing_Name = name_match.group(1).strip() 
        
        # Compare drawing names from list_PDF and drawing names from cells in Excel
        excel_Match = False # Flag to track if the drawing name matched in Excel
        # for name_cell in worksheet.iter_rows(min_row=24, max_row=150, min_col=3, max_col=3):        
        for row_index in range(24, 151):
            name_cell = worksheet.cell(row=row_index, column=3)
            # If there is a match, update revision in the same row at the revision column
            if name_cell.value and str(name_cell.value).strip() == drawing_Name:
                rev_cell = worksheet.cell(row=name_cell.row, column=rev_column)  # Revision is updated at this cell coordinate
                rev_cell.value = revision
                print(f"Matched '{drawing_Name}'. Updated revision to '{revision}' at row {name_cell.row}.")
                excel_Match = True
                break  # Exit the loop for this drawing file because a match was found
            
        # If not a match, add new drawing name into new row
        if not excel_Match :
            print(f"No match found for drawing: {drawing_Name}. Adding new entry...")
            # Find the next empty row in column C to add the new drawing
            for empty_row in range(24, 151):
                if worksheet.cell(row=empty_row, column=2).value is None:
                    worksheet.cell(row=empty_row, column=1, value=project_No)  # Add project number
                    worksheet.cell(row=empty_row, column=2, value=drawing_No)  # Add drawing number
                    worksheet.cell(row=empty_row, column=3, value=drawing_Name)  # Add drawing name
                    worksheet.cell(row=empty_row, column=rev_column, value=revision)  # Set initial revision to 1
                    print(f"Added new drawing {drawing_Name} at row {empty_row} with revision {revision}.")                            
                    break  # Exit the loop after adding the new drawing to avoid multiple additions

    # Extract transmittal file name from path
    pattern = r"transmittal \d{6}\.xlsx"
    match = re.search(pattern, transmittal, re.IGNORECASE)
    if match:
        transmittal_filename = match.group()
    else:
        print("Transmittal filename does not match expected pattern.")
        print("Please check the name matches: Transmittal YYMMDD.xlsx")
        return
    
    # Save the modified workbook    
    workbook.save(transmittal_filename)
    print(f"\nChanges saved in Transmittal excel '{transmittal_filename}'.")

    return transmittal

def Save_as_PDF():
    """
    Saves a specific worksheet from an Excel file to a PDF with A4 format
    by controlling the Excel application via COM.
    """
    excel_path = Overwrite_Transmittal()
    sheet_name = "CIVIL"

    currentDir = os.path.abspath(".")
    pattern = r"transmittal (.*)\.xlsx"
    match = re.search(pattern, excel_path, re.IGNORECASE)
    if match:
        pdf_name = f"Transmittal {match.group(1)}.pdf"
        pdf_path = os.path.join(currentDir, pdf_name)
    else:
        print("Transmittal filename does not match expected pattern.")
        print("Please check the name matches: Transmittal YYMMDD.xlsx")
        return

    print(f"\nConverting '{sheet_name}' from '{excel_path}' to PDF using MS Excel...")

    excel = None  # Initialize excel variable
    try:
        # Get absolute paths, which are required for COM
        excel_abs_path = os.path.abspath(excel_path)
        pdf_abs_path = os.path.abspath(pdf_path)

        # Start an instance of Excel
        excel = win32com.client.Dispatch("Excel.Application")
        # Keep the application hidden
        excel.Visible = False

        # Open the workbook
        workbook = excel.Workbooks.Open(excel_abs_path)

        # Select the specific worksheet
        worksheet = workbook.Worksheets[sheet_name]

        # --- Set Page Setup ---
        # xlPaperA4 has a value of 9
        worksheet.PageSetup.PaperSize = 9 
        # Default margins are used automatically.

        # --- Export to PDF ---
        # xlTypePDF has a value of 0
        worksheet.ExportAsFixedFormat(0, pdf_abs_path)
        
        print(f"Successfully saved PDF to '{pdf_abs_path}'")

    except Exception as e:
        print(f"An error occurred during PDF conversion: {e}")
    finally:
        # CRITICAL: Always ensure Excel is closed properly
        if excel:
            if 'workbook' in locals() and workbook:
                workbook.Close(SaveChanges=False)
            excel.Quit()
            # Release the COM object
            del excel

if __name__ == "__main__":    
    print("Transmit_Auto1000 Start")    
    Save_as_PDF()
    print("Created by Edd Palencia-Vanegas - June 2024. All rights reserved.")
    print("Version 5.0 - 21/10/2025")