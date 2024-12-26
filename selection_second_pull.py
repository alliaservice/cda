import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = "202402"

past_cda = pd.read_excel(os.path.join(abs_pth,f"all_titles_purchased_not_purchased.xlsx"))
first_p = pd.read_excel(os.path.join(abs_pth,f"{term}_selection/ds_first_all_titles.xlsx"))
second_p = pd.read_excel(os.path.join(abs_pth,f"{term}_selection/second_all_titles.xlsx"))

print("first pull: ", len(first_p))
print("second pull: ",len(second_p))
print("difference: ", (len(second_p) - len(first_p)))

# clean up the first pull with remove titles file
first_p['flag'] = 'y' # add simple flag column to use to exclude after merge
first_p = first_p[["ISBN", "flag"]] # keep only a few cols
first_p = first_p.drop_duplicates(subset=['ISBN']) # remove dupes on ISBN so merge works as expected


# merge first_p and second_p, keep only second_pull 

df_main = second_p.merge(first_p,
                     on = "ISBN",
                     how = 'left')

print("len after first/second merge: ", len(df_main))

# select only non_matching rows (ie second pull only titles)
df_main = df_main[df_main["flag"].isnull()]
print("len remove first: ", len(df_main))

# merge with past terms cda
df_main = df_main.merge(past_cda, 
                        on = "ISBN",
                        how = 'left')
print("past cda merge: ", len(df_main))
print(df_main.describe(include="all"))

df_main.to_excel(f'{abs_pth}/{term}_selection/{term}_second_pull_output.xlsx')
