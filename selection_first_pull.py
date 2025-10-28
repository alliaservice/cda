import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = "202502"

past_cda = pd.read_excel(os.path.join(abs_pth,f"all_titles_purchased_not_purchased.xlsx"))
first_p = pd.read_excel(os.path.join(abs_pth,f"{term}_selection/ds_first_all_titles.xlsx"))

# prevent duplicates in past cda
past_cda = past_cda.sort_values(["Purchased?", "Term_cda"], ascending=False) # first sort so most recent term is first
#test=past_cda[["Term_cda", "Purchased?"]]
#test.to_excel(f'{abs_pth}/{term}_selection/test1.xlsx')
#past_cda = past_cda.sort_values("Purchased?", ascending=False) # next sort so purchased is first
past_cda = past_cda.drop_duplicates(subset=['ISBN']) # remove dupes on ISBN so if there is a purchased title
# or there is a most recent title, only that one is kept and there aren't dupes
#print(past_cda["Purchased?"])

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
df_main.to_excel(f'{abs_pth}/{term}_selection/{term}_first_output-sort-test-again.xlsx')