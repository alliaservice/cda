import os
import pandas as pd
from openpyxl import Workbook
import sys

# define variables
term = "202402"
abs_pth = os.path.dirname(os.path.abspath(__file__))
cols_to_keep = ['Term','instructor_email', 'ISBN', 'CRN', 'Title_x', 
                'Price', 'Internal ID', 'LibSearch Link','DRM', 'Purchased?', #'removed',
                  'Term_cda']
output_directory = f'notification_{term}/reporting_output'

# read data
cda = pd.read_excel(os.path.join(abs_pth,f"notification_{term}/{term}_merge_matching.xlsx"))
#cda = cda.rename(columns={'flag': 'removed'}) # change name of flag to remove
cda = cda.rename(columns={"Instructor Email": "instructor_email"}) # rename to match
#print(cda['ISBN'])
def create_dir(abs_path:str, output_dir:str):
    '''
    abs_pth = the file path to the main directory scirpt is in, should be defined as global var
    output_dir = path from abs_path if there are intermediate directories and the output directory name
    creates a directory if none exists, if directory does exist, asks user for y/n input 
    if y, passes does not overwrite (but output files will be overwritten) 
    '''
    output_path = os.path.join(abs_path, output_dir)
    if os.path.exists(output_path):
        x = input('output directory exists, files will be overwritten. Continue? y/n')
        if x == 'n':
            sys.exit('you chose to quit')
        else: 
            pass

    else: 
        os.mkdir(output_path)

# create a directory for output files 
create_dir(abs_pth, output_directory)

# slice df, remove extra columns
#print(cda.keys()) # prints all column names for error checking
cda = cda[cols_to_keep] # edit cols_to_keep at the top to add or remove columns or change names

# add freq columns
# freq function
def add_frq(df, col:str):
    '''
    df: dataframe you want to add a freq column to
    col: column to base freq count on -- will also be new col 
    name ie CRN_freq
    adds a count column based on anther column
    '''
    df[f'{col}_freq'] = df.groupby(col)[col].transform('count')
    return df
# call freq function 3 times 
cda = add_frq(cda, 'CRN')
cda = add_frq(cda, 'instructor_email')
cda = add_frq(cda, 'ISBN')
#print(cda.head())

# save as an excel file

cda.to_excel(f'{abs_pth}/notification_{term}/reporting_output/{term}_output.xlsx')