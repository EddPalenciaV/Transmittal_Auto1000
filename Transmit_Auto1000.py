import os
import re
import shutil
import glob

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

if __name__ == "__main__":    
    print("Transmit_Auto1000 Start")