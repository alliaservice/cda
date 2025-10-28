import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
#term = "test"

lib_spending = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/fy24_25_cda_spending_raw.xlsx")
all_cda = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/all_cda_202301-202404.xlsx")

print(all_cda.keys())

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
                     how = 'outer',
                     indicator=True)

print(df_main.head(25))

#df_main.to_excel('lib_spending/output.xlsx')
df_main.to_excel('/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/all_cda_clean_25.xlsx')