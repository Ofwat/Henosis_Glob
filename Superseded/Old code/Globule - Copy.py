import pandas as pd
import os
import openpyxl
import glob
from datetime import datetime

path = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\1_Files as submitted by companies\Core"
path2 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\2_Files after Panacea\Core"
path3 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\3_Files error fixing WORKING\Core"

os.chdir(path3)  # Change the directory to the O drive

for file in glob.glob(r"*.xls*"):  # glob is a way of getting all the files of a certain type
    acronym = file[0:3]  # Get the three letter prefix we have added. This assumes that each file has the acronym of the company added at the beginning
    print(acronym)
    xl = pd.ExcelFile(file)  # Define the excel file
    worksheets = xl.sheet_names  # Get the list of worksheets in the file
    for i in worksheets:  # For each worksheet...
        if "fOut" in i:  # ...if the workseet contains the text "fOut" then...
            print(acronym + "_" + i)
            df_temp = pd.read_excel(file, sheet_name=i, skiprows=1)  # Put the data into a temporary dataframe
            df_temp.to_excel(".\\Disaggregated outputs\\" + acronym + "_" + i + ".xlsx", sheet_name="F_Outputs", startrow=1, index=False)

print("Adding text into cell B2")
for file in glob.glob(r".\Disaggregated outputs\*.xls*"):  # glob is a way of getting all the files of a certain type
    print(file)
    workbook = openpyxl.load_workbook(file)
    sheet = workbook['F_Outputs']
    now = datetime.now()
    sheet.cell(row=1, column=2).value = 'Python Text: ' + now.strftime("%H:%M")
    workbook.save(file)