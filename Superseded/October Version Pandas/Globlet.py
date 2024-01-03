
# THIS FILE HAS BEEN CREATED TO ENABLE ANY SUBSEQUENT ANALYSIS ON THE DATA
import pandas as pd
import os
import glob
import openpyxl
from datetime import datetime

path = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\1_Files as submitted by companies\Core"
path2 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\2_Files after Panacea\Core"
path3 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\3_Files error fixing WORKING\Core"

os.chdir(path3)  # Change the directory to the O drive

# Import from pickle file
df = pd.read_pickle(r".\Aggregated outputs\All.pkl")
companies = df["Acronym"].unique()  # Get the list of company acronyms

# Separate out the companies
for i in companies:
    print(i)
    df_temp = df[df["Acronym"] == i]
    # TODO: Delete the sheet_name column if not used for upload
    df_temp.to_excel(".\\Sliced outputs\\" + i + "_globlet.xlsx", sheet_name="F_Outputs", startrow=1, index=False)



# Takes around 3 mins to run

print("Adding text into cell B2")
for file in glob.glob(r".\Sliced outputs\*.xls*"):  # glob is a way of getting all the files of a certain type
    print(file)
    workbook = openpyxl.load_workbook(file)
    sheet = workbook['F_Outputs']
    now = datetime.now()
    sheet.cell(row=1, column=2).value = 'Python Text: ' + now.strftime("%H:%M")
    workbook.save(file)









