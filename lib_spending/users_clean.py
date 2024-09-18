import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = "test"

users = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/2023-07_20203-05_user_stats.xlsx")
all_cda = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/all_cda_dedupe.xlsx")


#users['ISBN'] = users['ISBN'].str.replace('-', '')
#slice = users.loc[['ISBN']]
#print(users['ISBN'])

#print(type(users['ISBN']))
all_cda['Title'] = all_cda['Title'].apply(lambda x:x.lower() if isinstance(x, str) else x)
users['Title'] = users['Title'].apply(lambda x:x.lower() if isinstance(x, str) else x)

df_main = users.merge(all_cda,
                     on = 'ISBN',
                     how = 'left',
                     indicator=True)
df_main = df_main.rename(columns={'_merge':'merge_isbn'})

df_main = df_main.merge(all_cda,
                        left_on = 'Title_x',
                        right_on = 'Title',
                        how = 'left',
                        indicator=True)

df_main.to_excel('lib_spending/users_output_title_isbn.xlsx')