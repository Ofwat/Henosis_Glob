

import pandas as pd
import os

df_mapping = pd.read_excel(r"C:\Users\Alex.whitmarsh\OneDrive - OFWAT\Data\Business Plan Tables\Version 6 Planning\231009 - Table to fOut mapping.xlsx",
                           sheet_name="BON mapping")

DWI_Tables = ["CW3", "OUT1", "OUT2", "OUT3", "OUT4", "OUT8", "OUT10"]
df_mapping2 = df_mapping[df_mapping["Table"].isin(DWI_Tables)]
df_mapping2 = df_mapping2[["Table", "BON"]]  # Just keep the relevant columns


# Alternative method which seems to give the same result
#pattern1 = '|'.join(DWI_Tables)
#df_mapping2 = df_mapping[df_mapping["Table"].str.contains(pattern1)]


print(df_mapping)


# Import the data
path2 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\2_Files after Panacea\Core"
os.chdir(path2)  # Change the directory to the O drive
df = pd.read_pickle(r".\Aggregated outputs\All.pkl")

df_dwi = df[df["Reference"].isin(df_mapping2["BON"])]

df_mapping2.to_excel("./Analysis for DWI/df_mapping2.xlsx")
df_dwi.to_excel("./Analysis for DWI/DWI numbers.xlsx")

print(df_dwi)
