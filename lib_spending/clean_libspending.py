import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = "test"

lib_spending = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/CDA_Expenditures_07192024-05072024.xlsx")
all_cda = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/all_cda_202301-03.xlsx")


all_cda['Title'] = all_cda['Title'].apply(lambda x:x.lower() if isinstance(x, str) else x)
all_cda['Title'] = all_cda['Title'].str.replace(' : ', ' ')
all_cda['Title'] = all_cda['Title'].str.replace(':','')
all_cda['Title'] = all_cda['Title'].str.replace("'", "")
all_cda['Title'] = all_cda['Title'].str.replace('-', ' ')
all_cda['Title'] = all_cda['Title'].str.replace('+', ' ')
all_cda['Title'] = all_cda['Title'].str.replace('/', '')
all_cda['Title'] = all_cda['Title'].str.replace(',','')
all_cda['Title'] = all_cda['Title'].str.replace('!', '')
all_cda['Title'] = all_cda['Title'].str.replace('?', '')
all_cda['Title'] = all_cda['Title'].str.replace('.', '')
all_cda['Title'] = all_cda['Title'].str.replace('"', '')
all_cda['Title'] = all_cda['Title'].str.replace('*', '')
all_cda['Title'] = all_cda['Title'].str.replace('%', '')
# would need to remove leading thes

df_main = lib_spending.merge(all_cda,
                     on = 'Title',
                     how = 'left',
                     indicator=True)

print(df_main.head(25))

#df_main.to_excel('lib_spending/output.xlsx')
all_cda.to_excel('all_cda_clean.xlsx')