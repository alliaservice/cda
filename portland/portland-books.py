import os
import pandas as pd
from openpyxl import Workbook

# define variables and read data 
abs_pth = os.path.dirname(os.path.abspath(__file__))
term = 202501 # CHANGE TO AN INPUT FIELD LATER
#term = input('enter the current term code YYYYTT for example 202501: ')

def get_excel_file(prompt):
    while True:
        file_name = input(prompt)
        file_path = os.path.join(abs_pth, file_name)
        try:
            df = pd.read_excel(file_path)
            return df
        except FileNotFoundError:
            print(f"File '{file_name}' not found. Please try again.")
        except Exception as e:
            print(f"Error: {e}. Please try again.")

#portland = get_excel_file('Enter Portland Courses file name:\nmust be an excel file ending in .xlsx for example 202501_portland.xlsx : ')
#print(portland)
#ds = get_excel_file('Enter Duck Store Book list file path: ')
#cda = get_excel_file('Enter CDA reporting file path: ')

#portland = pd.read_excel(os.path.join(abs_pth,input('Enter Portland Courses file name: \nmust be an excel ' \
#file ending in .xlsx for example 202501_portland.xlsx :')))
#print(portland)
#ds = pd.read_excel(os.path.join(abs_pth,input('Enter Duck Store Book list file path:')))
#cda = pd.read_excel(os.path.join(abs_pth,input('Enter CDA reporting file path:')))
portland = pd.read_excel(os.path.join(abs_pth,f"202501_portland.xlsx"))
ds = pd.read_excel(os.path.join(abs_pth,f"202501_full_ds.xlsx"))
cda = pd.read_excel(os.path.join(abs_pth,f"202501_output.xlsx"))

def concat(new_col:str, df, text1:str, col1:str, text2:str, text3:int):
    '''
    intended for ds links, concatenates text values and a column in a df.

    new_col = name of new column
    df = dataframe
    col1 = name of first column to concatenate, must contain strings or ints
    col2 = name of second column to concatenate, must contain strings or ints
    join_ch = character to join on, ie ';' or ' '
    '''
    df[new_col] = text1 + df[col1].astype(str) + text2 + str(text3)
# add duck store link
concat('duck_store_link', ds, 'https://www.uoduckstore.com/book-search-results?crn=', 'CRN','&term=',term)

# remove most cols from portland and cda
portland_strip = portland[['CAMPUS_DESC', 'SUBJECT', 'COURSE' ,'CRN', 'TITLE', 'PUBLISH_SUBJ_CODES', 'STATUS', 'PROGRAM', 'ACTUAL_ENROLLMENT']]
portland_strip = portland_strip.drop_duplicates(subset=['CRN']) # remove dupes on CRNs
cda = cda[['Internal ID', 'DRM', 'LibSearch Link', 'instructor_email']] # don't keep purchased -- not meaningful to pdx

# merge portland and ds booklist

merged = portland_strip.merge(ds,
                             on= 'CRN', 
                             how='left')

print('Number of rows in Portland course list, before deduplicating on CRN: ', len(portland))
print('Number of course sections in Portland course list, after deuplicating on CRN: ', len(portland_strip))
print('Number of rows in deduped Portland course list merged with book list: ', len(merged))  
print('Number of Portland course sections with books (or no materials) reported to the Duck Store:', merged['Internal ID'].count())
print('Number of books or free materials reported for Portland courses: ', merged['Title'].count())
print('Number of Portland courses reported "No Course Materials Required": ', merged['Req'].count("No Course Materials Required"))
print('Number of Portland courses that did not report to the Duck Store: ', merged['Internal ID'].isna().sum())
print('Numer of unique books that match Portland courses: ', merged['ISBN'].nunique())

# merge cda to merged
merged = merged.merge(cda,
                             on= 'Internal ID', 
                             how='left')
#print('merge with cda len: ', len(merged))
print('Number of books available from the library for Portland coureses: ', merged['DRM'].notna().sum())

# try with dupes 
merge_dupes = portland.merge(ds,
                             on= 'CRN', 
                             how='left')

merge_dupes = merge_dupes.merge(cda,
                             on= 'Internal ID', 
                             how='left')

print('\n Non de-duplicated:\n','Number of rows in un-deduped Portland course list merged with book list: ', len(merge_dupes))
print('Number of rows in Portland course list merged with book list: ', len(merge_dupes))  
print('Number of rows of books that match Portland courses:', merge_dupes['Internal ID'].count())
print('Number of rows of Portland courses with no books reported: ', merge_dupes['Internal ID'].isna().sum())
print('Number of rows of courses of books available from the library for Portland coureses: ', merge_dupes['DRM'].notna().sum())

# save files
# PROBABLY ADD A FAIL SAFE TO CHECK IF FILE EXISTS, also include term in file name
merged.to_excel(f'{abs_pth}/deduped_output.xlsx')
merge_dupes.to_excel(f'{abs_pth}/dupes_output.xlsx')   

# print some stats
