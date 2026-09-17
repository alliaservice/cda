import os
import pandas as pd
#from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = "202506"

# read in data
past_cda = pd.read_excel(os.path.join(abs_pth,f"all_titles_purchased_not_purchased.xlsx")) # list of all titles used in cda 
first_p = pd.read_excel(os.path.join(abs_pth,f"{term}_selection/ds_first_all_titles.xlsx")) # current term book list

# define column names (UPDATE HERE)
ISBN = 'ISBN' # in booklist file and past cda file
CRN = 'CRN' # in booklist file
purchased = 'Purchased?' # in past cda file, controlled vocab: no, owned/access, purchased
term_cda = 'Term_cda' # in past cda file, term purchased or first used in cda program
title_cda = 'Title_cda' # in past cda file

# prevent duplicates in past cda
past_cda = past_cda.sort_values([purchased, term_cda], ascending=False) # first sort so most recent term is first
past_cda = past_cda.drop_duplicates(subset=[ISBN]) # remove dupes on ISBN so if there is a purchased title
# or there is a most recent title, only that one is kept and there aren't dupes

#print(past_cda.keys(), first_p.keys()) # print column names of all input files.

print("first pull: ", len(first_p))

# merge with past terms cda
df_main = first_p.merge(past_cda, 
                        on = ISBN,
                        how = 'left')

# print stats
print("past cda merge: ", len(df_main))
#print(df_main.describe(include="all"))
non_matching = df_main[title_cda].isna().sum() # add up null values in title_cda
print('non-matching titles:',non_matching )
print('matching past terms', (len(df_main)-non_matching))


# add ds link, must remove .0 and convert crn to int
df_main[CRN] = df_main[CRN].round(0).astype('Int64') # convert crn to int for link
df_main["Duck Store Link"]= "https://www.uoduckstore.com/book-search-results?crn=" + df_main[CRN].astype(str) + f"&term={term}"

# save file 
df_main.to_excel(f'{abs_pth}/{term}_selection/{term}_first_output.xlsx')