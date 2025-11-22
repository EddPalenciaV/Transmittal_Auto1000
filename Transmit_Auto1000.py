import os
import re
import shutil

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


if __name__ == "__main__":    
    print("Transmit_Auto1000 Start")