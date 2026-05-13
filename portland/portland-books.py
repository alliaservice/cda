"""
Portland Books Filter Script

This script filters the Duck Store book list to only include books for
Portland campus courses. It also opptionally merges in CDA library titles.

Input files required:
files must be stored in the same folder as this script. 
File names need to follow this format, and include .xlsx extension.
term is the current term code in YYYYTT format, for example 202502 for Winter 2025.
    - {term}_portland.xlsx: Portland banner report (col names should be the same as example)
    - {term}_full_ds.xlsx: Duck Store book list (get from link to public folder on CDA sharepoint page)
    - {term}_output.xlsx: current term cda list (request from allia)

Output files:
    - {term}_deduped_output.xlsx: Processed data with deduplicated courses
    - {term}_dupes_output.xlsx: Processed data with all course sections (pdx banner report contains dupes)

Author: Allia Service
Email: aservice@uoregon.edu
Created: 09_2025
Last Modified: 10_2025

Dependencies:
    - pandas
    - openpyxl

Usage:
    python portland-books.py

"""

import os
import pandas as pd
from typing import Tuple, Optional
from openpyxl import Workbook

# define constants
required_file_message = """
The following files are required:
Files must be stored in the same folder as this script.
File names need to follow this format, and include .xlsx extension.
Term is the current term code in YYYYTT format, for example 202502 for Winter 2025.
- {term}_portland.xlsx: Portland banner report
- {term}_full_ds.xlsx: Duck Store book list
- {term}_output.xlsx: current term cda list
"""

def load_dataframes(term: str, abs_pth: str) -> Tuple[Optional[pd.DataFrame], 
Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Load all required dataframes for processing.
    term is the user input term (YYYYTT)
    abs_pth is the absolute path to the folder containing the files (defined in main)
    """
    portland = test_file(f"{term}_portland.xlsx", abs_pth)
    ds = test_file(f"{term}_full_ds.xlsx", abs_pth)
    cda = test_file(f"{term}_output.xlsx", abs_pth)
    return portland, ds, cda

def get_excel_file(prompt:str) -> pd.DataFrame:
    '''
    asks user for an excel file name until it doesn't get
    a file not found error.
    returns datafram.
    prompt = the text shown to the user.
    '''
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

def test_file(file_name, absolute_path) -> Optional[pd.DataFrame]:
    '''
    makes sure files exist and then reads them in as dataframes.
    if file doesn't exists, prints message and returns None.
    use an if statement to check for None after calling function.
    file_name = name of file to open including ext
    absolute_path = path to folder containing file
    '''
    file_path = os.path.join(absolute_path, file_name)
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print(f"File '{file_name}' not found in path '{absolute_path}'. \
              Please check the file names and term and try again.")
        return None
     
def get_unique_filename(base_path) -> str:
    """
    If file doesn't yet exist returns full file path as str
    If file exists returns file path with a unique filename by adding _1, _2 etc.
    iterates through numbers until it finds a file that doesn't exist.
    base_path = full file path
    """
    # might be better to use try except
    if not os.path.exists(base_path):
        return base_path
    
    name, ext = os.path.splitext(base_path)
    counter = 1
    
    while os.path.exists(f"{name}_{counter}{ext}"):
        counter += 1
    
    return f"{name}_{counter}{ext}"

def concat(new_col:str, df, text1:str, col1:str, text2:str, text3:int) -> None:
    '''
    intended for ds links, concatenates text values and a column in a df.
    returns None, adds new column to df.

    new_col = name of new column
    df = dataframe
    col1 = name of first column to concatenate, must contain strings or ints
    col2 = name of second column to concatenate, must contain strings or ints
    join_ch = character to join on, ie ';' or ' '
    '''
    df[new_col] = text1 + df[col1].astype(str) + text2 + str(text3)
def save_output_files(merged: pd.DataFrame, merge_dupes: pd.DataFrame, 
                     term: str, abs_pth: str) -> Tuple[str, str]:
    """Save output files and return filenames.
    intended to be for saving both deduped file and file w dupes.
    merged = deduped dataframe
    merge_dupes = dataframe with dupes
    """
    deduped_name = get_unique_filename(f'{abs_pth}/{term}_deduped_output.xlsx')
    dupe_name = get_unique_filename(f'{abs_pth}/{term}_dupes_output.xlsx')
    
    merged.to_excel(deduped_name)
    merge_dupes.to_excel(dupe_name)
    
    return deduped_name, dupe_name

def validate_term() -> str:
    """
    Validates user input for term code.
    Returns a valid 6-digit term code starting with 2 and ending with 1-4.
    """
    while True: # function will keep asking until valid input
        try:
            term = input('Enter the current term code YYYYTT (example 202501): ')
            # Check if input is 6 digits
            if not term.isdigit() or len(term) != 6:
                raise ValueError("Term must be exactly 6 digits")
            # Check if starts with 2
            if not term.startswith('2'):
                raise ValueError("Term must start with 2")
            # Check if ends with 1-4
            if term[-1] not in '1234':
                raise ValueError("Term must end with 1, 2, 3, or 4")
            return term
        except ValueError as e:
            print(f"Invalid term: {e}")
            print("Please try again.")

def main():
    '''
    '''
    # define variables and initialize
    abs_pth = os.path.dirname(os.path.abspath(__file__))
    print(required_file_message) 
    term = validate_term()

    # Load data
    portland, ds, cda = load_dataframes(term, abs_pth)
    if any(df is None for df in [portland, ds, cda]):
        print("Missing input files. Please check filenames and try again.")
        return
    # check for missing ds cols that we'll need later
    try:
        test = ds[['CRN', 'Internal ID', 'Title', 'ISBN', 'Req']]
    except KeyError as e: # error handling for missing/renamed cols
        print(f"Error: Missing expected column in Duck Store file: {e}",
              "\n please check column names and rerun this program.")
    
    # add duck store link column
    concat('duck_store_link', ds, 'https://www.uoduckstore.com/book-search-results?crn=',
            'CRN','&term=',term)
    
    # clean up dfs: remove most cols from portland and cda, dedupe portland on CRN
    try:
        portland_strip = portland[['CAMPUS_DESC', 'SUBJECT', 'COURSE' ,'CRN', 'TITLE', 
                                'PUBLISH_SUBJ_CODES', 'STATUS', 'PROGRAM', 'ACTUAL_ENROLLMENT']]
        portland_strip = portland_strip.drop_duplicates(subset=['CRN']) # remove dupes on CRNs
    except KeyError as e: # error handling for missing/renamed cols
        print(f"Error: Missing expected column in Portland banner file: {e}",
              "\n please check column names and rerun this program.")
    
    try:
        cda = cda[['Internal ID', 'DRM', 'LibSearch Link', 
               'instructor_email']] # don't keep purchased -- not meaningful to pdx
    except KeyError as e:
        print(f"Error: Missing expected column in CDA file: {e}",
              "\n please check column names and rerun this program.")

    # merge portland and ds booklist

    merged = portland_strip.merge(ds, on= 'CRN', how='left')

    print('Number of rows in Portland course list, before deduplicating on CRN: ', len(portland))
    print('Number of course sections in Portland course list, after deuplicating on CRN: ', len(portland_strip))
    print('Number of rows in deduped Portland course list merged with book list: ', len(merged))  
    print('Number of Portland course sections with books (or no materials) reported to the Duck Store:', merged['Internal ID'].count())
    print('Number of books or free materials reported for Portland courses: ', merged['Title'].count())
    print('Number of Portland courses reported "No Course Materials Required": '
          , merged['Req'].eq("No Course Materials Required").sum()) #.eq returns boolean series comparing each col value to string
    print('Number of Portland courses that did not report to the Duck Store: ', merged['Internal ID'].isna().sum())
    print('Numer of unique books that match Portland courses: ', merged['ISBN'].nunique())

    # merge cda to merged
    merged = merged.merge(cda,on= 'Internal ID', how='left')
    #print('merge with cda len: ', len(merged))
    print('Number of books available from the library for Portland coureses: ', merged['DRM'].notna().sum())

    # merge without deduping portland (use portland not portland_strip)
    merge_dupes = portland.merge(ds,on= 'CRN',  how='left')
    merge_dupes = merge_dupes.merge(cda, on= 'Internal ID', how='left')

    print('\n Non de-duplicated:\n','Number of rows in un-deduped Portland course list merged with book list: ', len(merge_dupes))
    print('Number of rows in Portland course list merged with book list: ', len(merge_dupes))  
    print('Number of rows of books that match Portland courses:', merge_dupes['Internal ID'].count())
    print('Number of rows of Portland courses with no books reported: ', merge_dupes['Internal ID'].isna().sum())
    print('Number of rows of courses of books available from the library for Portland coureses: ', merge_dupes['DRM'].notna().sum())

    # save files, if files exist, add _1, _2 etc to filename
    deduped_name, dupe_name = save_output_files(merged, merge_dupes, term, abs_pth)
    print(f'\nFiles saved successfully as {deduped_name} and {dupe_name}.')

    
        



if __name__ == "__main__":
    main()





# print some stats



# not using
#portland = get_excel_file('Enter Portland Courses file name:\nmust be an excel file ending in .xlsx for example "202501_portland.xlsx" : ')
#print(portland)
#ds = get_excel_file('Enter Duck Store Book list file path: ')
#cda = get_excel_file('Enter CDA reporting file path: ')

#portland = pd.read_excel(os.path.join(abs_pth,input('Enter Portland Courses file name: \nmust be an excel ' \
#file ending in .xlsx for example 202501_portland.xlsx :')))
#print(portland)
#ds = pd.read_excel(os.path.join(abs_pth,input('Enter Duck Store Book list file path:')))
#cda = pd.read_excel(os.path.join(abs_pth,input('Enter CDA reporting file path:')))

# dedupe


def print_stats(df: pd.DataFrame, type: str) -> None:
    '''
    print stats for the merged dataframes.
    some stats will be printed seperately in the main function.
    '''
    print(f'\n {type} stats:')  
    print('Number of rows Portland course list merged with book list: ', len(df))  
    print('Number of Portland course sections that reported (could report materials or no materials) ' \
    'reported to the Duck Store:', df['Internal ID'].count()) # counts rows with non blank ds rows
    print('Number of Portland courses that did not report to the Duck Store: ', df['Internal ID'].isna().sum())
    print('Number of books or free materials reported for Portland courses' \
    '(nubmer of rows with a book store "Title" listed): ', df['Title'].count())
    print('Number of Portland courses that reported "No Course Materials Required": '
        , df['Req'].eq("No Course Materials Required").sum()) #.eq returns boolean series comparing each col value to string
    print('Number of unique books used in Portland courses (ISBNs): ', df['ISBN'].nunique())
    print('Number of rows of courses of books available from the library for Portland coureses: ', df['DRM'].notna().sum())





