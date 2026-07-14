import openpyxl
import glob
import os
import openpyxl.writer
import openpyxl.writer.excel
import pandas as pd
from pathlib import Path
import shutil
import time
import warnings
import logging 


def mySleepFunction(seconds):
    for i in range(seconds):
        print(f"Waiting... {seconds - i} seconds remaining", end="\r")
        time.sleep(1)

#Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
    return rangeSelected

#Takes: start cell, end cell, and sheet you want to copy from.
def copyRangeInternalValue(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    #Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
        #Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).internal_value)
        #Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
    return rangeSelected

#Paste range
#Paste data from copyRange into template sheet
def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData):
    countRow = 0
    for i in range(startRow,endRow+1,1):
        countCol = 0
        for j in range(startCol,endCol+1,1):
            sheetReceiving.cell(row=i, column=j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1

def Merge_All_files():  
    xlfiles = list(base_path.glob("WorkinProgress_*.xlsx"))
    if len(xlfiles) == 0:
        print(f"No files found in the directory: {base_path}")
        return
    
    print("Starting Optimized A2 Copy Merge Routine with Dynamic Headers...")
    template_write = openpyxl.load_workbook(report_file, data_only=False)
    temp_sheet = template_write["Sheet1"] 
    
    # Track if we have written the header row yet
    header_written = False
    current_write_row = 2 # Data starts at row 2

    for files_list in xlfiles:
        block_name = files_list.stem.split("_")[1]
        year = files_list.stem.split("_")[-1]
        
        print(f"Processing File -> Block: {block_name} | Year: {year}")
        
        wb = openpyxl.load_workbook(files_list, read_only=True) 
        sheet = wb['Sheet1']
        
        # 1. NEW: Extract and assign headers from the VERY FIRST file
        if not header_written:
            print("Extracting headers from the first source file...")
            # Grab exactly row 1 from the source file
            for first_row in sheet.iter_rows(min_row=1, max_row=1, values_only=True):
                temp_sheet.cell(row=1, column=1, value="Block Name")
                temp_sheet.cell(row=1, column=2, value="Financial Year")
                
                # Copy the rest of the original headers starting at Column 3
                for col_idx, header_value in enumerate(first_row, start=3):
                    temp_sheet.cell(row=1, column=col_idx, value=header_value)
            
            header_written = True # Lock it so we don't overwrite headers again
        
        # 2. Stream the data rows (skipping the header via min_row=2)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
                
            temp_sheet.cell(row=current_write_row, column=1, value=block_name)
            temp_sheet.cell(row=current_write_row, column=2, value=year)
            
            for col_idx, cell_value in enumerate(row, start=3):
                temp_sheet.cell(row=current_write_row, column=col_idx, value=cell_value)
            
            current_write_row += 1
            
        wb.close()  
        print(f"Successfully appended: {files_list.name}")
        
    print("Saving consolidated data...")
    template_write.save(report_file)
    mySleepFunction(2)
    template_write.close()
    print("All files and headers merged successfully!")


# excution starts here
##########################################################################################################
#dir_path = os.path.dirname(os.path.realpath(__file__)) 
base_path = Path("C:\\Users\\HP\\Downloads\\WIP")
base_file = "Workin_progress_All.xlsx"
report_file = base_path.joinpath(base_file)
backup_folder = base_path.joinpath("backup_folder")


print(base_path)
print(report_file)

warnings.simplefilter("ignore")
file_list_xlsx = list(Path(base_path).glob("*.xls"))

#Converting xls file into xlsx
for f in file_list_xlsx:
    print(f"Converting: {f}")
    
    # 1. Read the HTML table into a list of DataFrames
    # [0] gets the first table found in the file
    data_list = pd.read_html(f)
    df = data_list[0]
    
    # 2. Create the new filename (switching .xls to .xlsx)
    # Using Path (from pathlib) is cleaner than .replace()
    new_filename = f.with_suffix('.xlsx') 
    
    # 3. Save as a real Excel file
    df.to_excel(new_filename, index=False)
    
    # 4. Remove the old fake .xls file
    if os.path.isdir(backup_folder):
        print("folder exist, moving original files")
    else:
        Path(backup_folder).mkdir(parents=True, exist_ok=True)
        
    shutil.move(f, base_path.joinpath("backup_folder"))  # Move the original .xls to a backup folder instead of deleting
    #os.remove(f)
    print(f"Converted and moved: {f}")
Merge_All_files()










    
    

