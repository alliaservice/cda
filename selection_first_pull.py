import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = "test"

past_cda = pd.read_excel(os.path.join(abs_pth,f"all_titles_purchased_not_purchased.xlsx"))
first_p = pd.read_excel(os.path.join(abs_pth,f"{term}_selection/ds_first_all_titles.xlsx"))
#print(past_cda.head())

print("first pull: ", len(first_p))

# merge with past terms cda
df_main = first_p.merge(past_cda, 
                        on = "ISBN",
                        how = 'left')

# print stats
print("past cda merge: ", len(df_main))
#print(df_main.describe(include="all"))
non_matching = df_main["Title_cda"].isna().sum() # add up null values in title_cda
print('non-matching titles:',non_matching )
print('matching past terms', (len(df_main)-non_matching))

# save file 
df_main.to_excel(f'{abs_pth}/{term}_selection/{term}_first_output.xlsx')