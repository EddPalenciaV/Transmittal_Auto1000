import os
import re
import shutil
import glob
from openpyxl import load_workbook

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

if __name__ == "__main__":    
    print("Transmit_Auto1000 Start")