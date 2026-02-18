import openpyxl
import glob
import os
import openpyxl.writer
import openpyxl.writer.excel
import pandas as pd
from pathlib import Path


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

            sheetReceiving.cell(row = i, column = j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1

##########################################################################################################
def Swap_A2_for_progress():
    print("Swaping Current Data to Old for progress...")
    
    wb = openpyxl.load_workbook(report_file) 
    New_sheet = wb['2024-26 A2']
    old_sheet = wb['2024-26 A2_old']
    copiedData=copyRange(3, 2, 15, 484, New_sheet)
    pasteRange(3, 2, 15, 484, old_sheet, copiedData)

def Swap_A4_for_progress():
    print("Swaping Current A4 Data to Old for progress...")
    
    wb = openpyxl.load_workbook(report_file) 
    New_sheet = wb['NewA4Report_2426']
    old_sheet = wb['OldA4Report_2426']
    copiedData=copyRange(3, 3, 29, 485, New_sheet)
    pasteRange(3, 3, 29, 485, old_sheet, copiedData)


##########################################################################################################
def CopyA2_2425_Data():
    xlfiles = glob.glob(RAW_FILE+ "Phy*2024*.xlsx")
    print(xlfiles)

    #File to be copied
    print("starting A2 Copy")
    #File to be pasted into
    template = openpyxl.load_workbook(report_file) 
    # Copying A2 Files 2024-25
    temp_sheet = template["2024-24 A2"] 
    for files_list in xlfiles:
        
        wb = openpyxl.load_workbook(files_list) 
        sheet = wb['Sheet1']
        
        #copyRange(startCol, startRow, endCol, endRow, sheet)
        #pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData)
        files = Path(files_list).name
        print("Copying: "+ files +" in main excel file")
        match files:
            #Bhaiyathan 24-25
            case "PhysicalProgressReport_PMAYG_3305012_2024-2025.xlsx":
                print("Copying Bhaiyathan 24-25")
                copiedData=copyRange(1, 3, 13, 80, sheet)
                pasteRange(3, 2, 15, 79, temp_sheet,copiedData)
                                
            #Odgi 24-25
            case "PhysicalProgressReport_PMAYG_3305013_2024-2025.xlsx":
                print("Copying Odgi 24-25")
                copiedData=copyRange(1, 3, 13, 76, sheet)
                pasteRange(3, 80, 15, 153, temp_sheet,copiedData)
                
            #Pratappur 24-25
            case "PhysicalProgressReport_PMAYG_3305015_2024-2025.xlsx":
                print("Copying Pratappur 24-25")
                copiedData=copyRange(1, 3, 13, 104, sheet)
                pasteRange(3, 154, 15, 255, temp_sheet,copiedData)
                
            #Premnagar 24-25
            case "PhysicalProgressReport_PMAYG_3305010_2024-2025.xlsx":
                print("Copying Premnagar 24-25")
                copiedData=copyRange(1, 3, 13, 49, sheet)
                pasteRange(3, 256, 15, 302, temp_sheet,copiedData)
                
            #Ramanujnagar 24-25
            case "PhysicalProgressReport_PMAYG_3305011_2024-2025.xlsx":
                print("Copying Ramanujnagar 24-25")
                copiedData=copyRange(1, 3, 13, 76, sheet)
                pasteRange(3, 303, 15, 376, temp_sheet,copiedData)
                
            #Surajpur 24-25
            case "PhysicalProgressReport_PMAYG_3305009_2024-2025.xlsx":
                print("Copying Surajpur 24-25")
                copiedData=copyRange(1, 3, 13, 110, sheet)
                pasteRange(3, 377, 15, 484, temp_sheet,copiedData)
                
            case _:
                print("file not matching with any case:" + files)

        wb.close()  
 
    print("A2 File saved")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()
        
    print("All files copied and pasted successfully")

##########################################################################################################
def CopyA4_2425_Data():
   
    xlfiles = glob.glob(RAW_FILE + "Gap*.xlsx")
    print(xlfiles)
 
    #File to be copied
    print("starting A4 Copy")
    #File to be pasted into
    template = openpyxl.load_workbook(report_file) #Add file name
 
    # Copying A4 Files 2024-25
    print("Opening A4 Sheet")
    temp_sheet = template["NewA4Report_2425"] 
    for files_list in xlfiles:
        
        wb = openpyxl.load_workbook(files_list) 
        sheet = wb['Sheet1']
        
        #copyRange(startCol, startRow, endCol, endRow, sheet)
        #pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData)
        files = Path(files_list).name
        print("Copying: "+ files +" in main excel file")
        match files:
            #Bhaiyathan 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305012_2024-2025.xlsx":
                print("Copying Bhaiyathan 24-25")
                copiedData=copyRange(1, 3, 27, 81, sheet)
                pasteRange(3, 3, 29, 80, temp_sheet,copiedData)
                                
            #Odgi 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305013_2024-2025.xlsx":
                print("Copying Odgi 24-25")
                copiedData=copyRange(1, 3, 27, 77, sheet)
                pasteRange(3, 81, 29, 154, temp_sheet,copiedData)
                
            #Pratappur 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305015_2024-2025.xlsx":
                print("Copying Pratappur 24-25")
                copiedData=copyRange(1, 3, 27, 105, sheet)
                pasteRange(3, 155, 29, 256, temp_sheet,copiedData)
                
            #Premnagar 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305010_2024-2025.xlsx":
                print("Copying Premnagar 24-25")
                copiedData=copyRange(1, 3, 27, 50, sheet)
                pasteRange(3, 257, 29, 303, temp_sheet,copiedData)
                
            #Ramanujnagar 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305011_2024-2025.xlsx":
                print("Copying Ramanujnagar 24-25")
                copiedData=copyRange(1, 3, 27, 77, sheet)
                pasteRange(3, 304, 29, 377, temp_sheet,copiedData)
                
            #Surajpur 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305009_2024-2025.xlsx":
                print("Copying Surajpur 24-25")
                copiedData=copyRange(1, 3, 27, 111, sheet)
                pasteRange(3, 378, 29, 485, temp_sheet,copiedData)
                
            case _:
                print("file not matching with any case:" + files)

        wb.close()  
        print("done")

    print("A4 File saved")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()
        
    print("All files copied and pasted successfully")

##########################################################################################################
def CopyA2_2425_Data_old():

    print("Swaping today's A2 To Old Data")
    #File to be pasted into
    print(report_file)
    template = openpyxl.load_workbook(report_file) #Add file name
    # Copying A2 Files 2024-25
    a2_new = template["2024-24 A2"]
    a2_old = template["2024-24 A2_old"]
    copied_data = copyRange(3, 2, 15, 484, a2_new)
    pasteRange(3, 2, 15, 484, a2_old,copied_data)
    print("Swaping A2 File completed.")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()


##########################################################################################################
def CopyA4_2425_Data_old():
    print("Swaping today's A4 To Old Data")
    #File to be pasted into
    template = openpyxl.load_workbook(report_file) #Add file name
    # Copying A2 Files 2024-25
    a2_new = template["NewA4Report_2425"]
    a2_old = template["OldA4Report_2425"]
    copied_data = copyRange(3, 3, 29, 485, a2_new)
    pasteRange(3, 3, 29, 485, a2_old,copied_data)
    print("Swaping A4 File completed.")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()

##########################################################################################################
def CopyA4_1625_Data_old():
    print("Swaping today's A4 To Old Data")
    #File to be pasted into
    template = openpyxl.load_workbook(report_file, data_only=True) #Add file name
    # Copying A2 Files 2024-25
    a2_new = template["NewA4Report16_23"]
    a2_old = template["OldA4Report16-23"]
    copied_data = copyRange(3, 3, 29, 485, a2_new)
    pasteRange(3, 3, 29, 485, a2_old,copied_data)
    print("Swaping 16-23 A4 in old File completed.")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()    

##########################################################################################################
def CopyA2_1625_Data():

    xlfiles = glob.glob(RAW_FILE+"Phy*ALL*.xlsx")
    print(xlfiles)

    #File to be copied
    print("starting A2 16-25 Copy")
    #File to be pasted into
    template = openpyxl.load_workbook(report_file) 
    # Copying A2 Files 1623
    temp_sheet = template["A2_Report_16_23"] 
    for files_list in xlfiles:
        
        wb = openpyxl.load_workbook(files_list) 
        sheet = wb['Sheet1']
        
        #copyRange(startCol, startRow, endCol, endRow, sheet)
        #pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData)
        files = Path(files_list).name
        print("Copying: "+ files +" in main excel file")
        match files:
            #Bhaiyathan 24-25
            case "PhysicalProgressReport_PMAYG_3305012_ALLPMAYG.xlsx":
                print("Copying Bhaiyathan 16-25")
                copiedData=copyRange(1, 3, 13, 80, sheet)
                pasteRange(31, 3, 43, 80, temp_sheet,copiedData)
                                
            #Odgi 24-25
            case "PhysicalProgressReport_PMAYG_3305013_ALLPMAYG.xlsx":
                print("Copying Odgi 16-25")
                copiedData=copyRange(1, 3, 13, 76, sheet)
                pasteRange(31, 81, 43, 154, temp_sheet,copiedData)
                
            #Pratappur 24-25
            case "PhysicalProgressReport_PMAYG_3305015_ALLPMAYG.xlsx":
                print("Copying Pratappur 16-25")
                copiedData=copyRange(1, 3, 13, 104, sheet)
                pasteRange(31, 155,43, 256, temp_sheet,copiedData)
                
            #Premnagar 24-25
            case "PhysicalProgressReport_PMAYG_3305010_ALLPMAYG.xlsx":
                print("Copying Premnagar 16-25")
                copiedData=copyRange(1, 3, 13, 49, sheet)
                pasteRange(31, 257,43, 303, temp_sheet,copiedData)
                
            #Ramanujnagar 24-25
            case "PhysicalProgressReport_PMAYG_3305011_ALLPMAYG.xlsx":
                print("Copying Ramanujnagar 16-25")
                copiedData=copyRange(1, 3, 13, 76, sheet)
                pasteRange(31, 304,43, 377, temp_sheet,copiedData)
                
            #Surajpur 24-25
            case "PhysicalProgressReport_PMAYG_3305009_ALLPMAYG.xlsx":
                print("Copying Surajpur 16-25")
                copiedData=copyRange(1, 3, 13, 110, sheet)
                pasteRange(31, 378,43, 485, temp_sheet,copiedData)
                
            case _:
                print("file not matching with any case:" + files)
  

        wb.close()  
 
    print("A2 File saved")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()
        
    print("All files copied and pasted successfully")

def CopyA4_1625_Data():
    xlfiles = glob.glob( RAW_FILE+"Gap*ALL*.xlsx")
    print(xlfiles)
 
    #File to be copied
    print("starting A4 16-25 Copy")
    #File to be pasted into
    template = openpyxl.load_workbook(report_file) #Add file name
 
    # Copying A4 Files 2024-25
    print("Opening A4 Sheet")
    temp_sheet = template["NewA4Report16_23"] 
    for files_list in xlfiles:
        
        wb = openpyxl.load_workbook(files_list) 
        sheet = wb['Sheet1']
        
        #copyRange(startCol, startRow, endCol, endRow, sheet)
        #pasteRange(startCol, startRow, endCol, endRow, sheetReceiving,copiedData)
        files = Path(files_list).name
        print("Copying: "+ files +" in main excel file")
        match files:
            #Bhaiyathan 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305012_ALLPMAYG.xlsx":
                print("Copying Bhaiyathan 24-25")
                copiedData=copyRange(1, 3, 27, 81, sheet)
                pasteRange(59, 3, 85, 80, temp_sheet,copiedData)
                                
            #Odgi 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305013_ALLPMAYG.xlsx":
                print("Copying Odgi 24-25")
                copiedData=copyRange(1, 3, 27, 77, sheet)
                pasteRange(59, 81, 85, 154, temp_sheet,copiedData)
                
            #Pratappur 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305015_ALLPMAYG.xlsx":
                print("Copying Pratappur 24-25")
                copiedData=copyRange(1, 3, 27, 105, sheet)
                pasteRange(59, 155,85, 256, temp_sheet,copiedData)
                
            #Premnagar 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305010_ALLPMAYG.xlsx":
                print("Copying Premnagar 24-25")
                copiedData=copyRange(1, 3, 27, 50, sheet)
                pasteRange(59, 257,85, 303, temp_sheet,copiedData)
                
            #Ramanujnagar 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305011_ALLPMAYG.xlsx":
                print("Copying Ramanujnagar 24-25")
                copiedData=copyRange(1, 3, 27, 77, sheet)
                pasteRange(59, 304,85, 377, temp_sheet,copiedData)
                
            #Surajpur 24-25
            case "GapinprogressAccountverifctioncompletion_PMAYG_3305009_ALLPMAYG.xlsx":
                print("Copying Surajpur 24-25")
                copiedData=copyRange(1, 3, 27, 111, sheet)
                pasteRange(59, 378,85, 485, temp_sheet,copiedData)
                
            case _:
                print("file not matching with any case:" + files)
 

        wb.close()  
        print("done")

    print("A4 File saved")
    openpyxl.writer.excel.save_workbook(template, report_file)
    template.close()
    print("All files copied and pasted successfully")

# excution starts here
##########################################################################################################
dir_path = os.path.dirname(os.path.realpath(__file__))
report_file = dir_path + "/PMAYG_TA-AM_WISE_REPORT_07012026_13022026_COLL_M.xlsx"
RAW_FILE = dir_path+ "/portalData/"
print(RAW_FILE)
file_list = glob.glob(RAW_FILE + "*.xls")

# Converting xls file into xlsx
# for f in file_list:
#     print(f)
#     data = pd.read_html(f)
#     data[0].to_excel(f.replace(".xls", ".xlsx"), index=False)
#     #os.remove(f)

## 24-25 Files
#swap_2425
Swap_A2_for_progress()
#CopyA2_2425_Data()
#CopyA2_2526_Data()

Swap_A4_for_progress()
#CopyA4_2425_Data()
#CopyA4_2526_Data()








    
    

