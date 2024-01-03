
# THIS FILE HAS BEEN CREATED TO ENABLE ANY SUBSEQUENT ANALYSIS ON THE DATA
import pandas as pd
import os

path = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\1_Files as submitted by companies\Core"
path2 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\2_Files after Panacea\Core"
path3 = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\3_Files error fixing WORKING\Core"

os.chdir(path3)  # Change the directory to the O drive

# Import from pickle file
df = pd.read_pickle(r".\Aggregated outputs\All.pkl")

# Fix 0x7
#df_fixed = df.replace("0x7", '')
#df_batch1.to_excel(r".\Aggregated outputs\Batch1.xlsx", sheet_name="F_Outputs")


# Sort the data according to BON code
#df_sort = df.sort_values(by='Reference')
#df_sort.reset_index(inplace=True, drop=True)
#df_sort.to_excel(r".\Aggregated outputs\All_sorted.xlsx", sheet_name="F_Outputs")

# Slicing by batches
batch1 = ["ANH", "SEW", "SSC", "SVE", "UUW"]  # minus SVE and SSC
batch2 = ["PRT", "SRN", "SBB", "WSX", "YKY"]
batch3 = ["AFW", "HDD", "NES", "SES", "TMS", "WSH"]
pattern1 = '|'.join(batch1)
#df_batch1 = df[df["Acronym"].str.contains(pattern1)]
#df_batch1.to_excel(r".\Aggregated outputs\Batch1.xlsx", sheet_name="F_Outputs")

df_globlet = df[df["Acronym"] == "SBB"]
df_globlet.to_excel(r".\Aggregated outputs\SBB_globlet.xlsx", sheet_name="F_Outputs")



# Duplicates analysis
#f_sheets = df['Sheet_name'].unique()
#df2 = df[['Acronym', 'Sheet_name']]
#df2.drop_duplicates(inplace=True)
#df2['Sheet Present'] = 'Yes'
#df2 = df2.pivot(index='Acronym', columns='Sheet_name', values='Sheet Present')


#duplicated_rows = df.duplicated(subset=['Acronym', 'Reference'], keep=False)
#df3 = df[duplicated_rows]
#df3.to_excel(r".\Aggregated outputs\Duplicated BON code data.xlsx", sheet_name="F_Outputs")

#df_first = df[df.duplicated(subset=['Acronym', 'Reference'], keep="first")]
#df_last = df[df.duplicated(subset=['Acronym', 'Reference'], keep="last")]
#df_diff = df_first.compare(df_last, align_axis=0)
#print(df_diff)

company_order = ["ANH", "WSH", "HDD", "NES", "SVE", "SBB", "SRN", "TMS", "UUW", "WSX", "YKY"
                 "AFW", "PRT", "SEW", "SSC", "SES"]










