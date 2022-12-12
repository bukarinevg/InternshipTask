import os
import pandas as pd
import openpyxl as op
from datetime import datetime, timedelta
from pandas.io.excel import ExcelWriter

# Getting data from files
os.chdir("C:\etl\ETL")
mapping_labs = pd.read_excel('./Mapping.xlsx', sheet_name='Labs')
mapping_drugs = pd.read_excel('./Mapping.xlsx', sheet_name='Drugs')
mapping_conditions = pd.read_excel('./Mapping.xlsx', sheet_name='Conditions')
general = pd.read_excel('./General.xlsx')
drugs = pd.read_excel('./Drugs.xlsx')
conditions = pd.read_excel('./Conditions.xlsx')
labs = pd.read_excel('./Labs.xlsx')

# Filtered drugs
drugs = drugs[(drugs['Claim_Status'] == 'APPR') & (pd.Timestamp.today() - timedelta(days=365) < drugs['Fill_Date'])]

# Put information about drugs
for i, row in general.iterrows():
    member_drugs = drugs[drugs['Member_ID'] == row['Member_ID']]
    member_drugs.to_excel('members/'+row['Full_Name'] +'.xlsx', sheet_name='Drugs', index=False)

# Put information about drugs
for i, row in general.iterrows():
    member_labs = labs[labs['Member_ID'] == row['Member_ID']]
    with ExcelWriter('./members/'+row['Full_Name'] +'.xlsx', mode="a" if os.path.exists('./members/'+row['Full_Name'] +'.xlsx') else "w") as writer:
        member_labs.to_excel(writer, sheet_name='Labs', index=False)


# Put information about conditions
for i, row in general.iterrows():
    member_conditions = conditions[conditions['Member_ID'] == row['Member_ID']]
    with ExcelWriter('./members/'+row['Full_Name'] +'.xlsx', mode="a" if os.path.exists('./members/'+row['Full_Name'] +'.xlsx') else "w") as writer:
        member_conditions.to_excel(writer, sheet_name='Conditions', index=False)
