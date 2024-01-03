import pandas as pd
import os
import openpyxl
import glob
from datetime import datetime

path = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\1_Files as submitted by companies\Core"
path2 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\2_Files after Panacea\Core"
path3 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\3_Files error fixing WORKING\Core"

os.chdir(path3)  # Change the directory to the O drive

df = pd.read_pickle(r".\Aggregated outputs\All.pkl")
df["Globule"] = df["Acronym"] + "_" + df["Sheet_name"]

fOuts = df["Globule"].unique()  # Get the list of company acronyms

# Separate out the fOUts
for i in fOuts:
    print(i)
    df_temp = df[df["Globule"] == i]
    df_temp = df_temp.drop(columns = ["Sheet_name", "Globule"])
    df_temp.to_excel(".\\Disaggregated outputs\\" + i + ".xlsx", sheet_name="F_Outputs", startrow=1, index=False)

print("Adding text into cell B2")
for file in glob.glob(r".\Disaggregated outputs\*.xls*"):  # glob is a way of getting all the files of a certain type
    print(file)
    workbook = openpyxl.load_workbook(file)
    sheet = workbook['F_Outputs']
    now = datetime.now()
    sheet.cell(row=1, column=2).value = 'Python Text: ' + now.strftime("%H:%M")
    workbook.save(file)