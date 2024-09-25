import pandas as pd
import os
import sys

# define variables
term = "202304"
term_int = int(term)
abs_pth = os.path.dirname(os.path.abspath(__file__))
output_dir = f'notification_{term}/script2_output'
output_dir_path = os.path.join(abs_pth, output_dir)

# read data
qc_matching = pd.read_excel(os.path.join(abs_pth,f"notification_{term}/{term}_merge_matching_clean.xlsx"))

# cols to keep for horizonal x_emails files, update email_cols as well for changes
cols_to_keep = ['instructor_email', 'Title_x', 'instructor_full_name', 
                'LibSearch Link','print_link?','print_libsearch_link', 'DRM',
                'Purchased?', 'course_numbers', 'license_text', 'CRN_freq', 
                'instructor_email_freq', 'CRN']
email_cols = ['Title_x', 'LibSearch Link', 'DRM',
                              'Purchased?', 'course_numbers', 'license_text','print_link?',
                              'print_libsearch_link', 'CRN_freq', 'CRN']

# vars for license text for emails, used in horizontal explode func, call to conditional col
limit_ex_search = [['cop', ("A limited license means that the library wasn't able to purchase unlimited " 
                   "access to the title. If too many students try to access the title at the same time, "
                   "some will be denied access. If we notice that lots of students are trying to access "
                   "the book, we'll attempt to buy additional copies, but it may take a few days. "
                   "So what does this mean for you and your students? In many cases students won't access "
                   "the book concurrently, so we'll be able to keep up for demand on the book, but if "
                   "students want guaranteed access at all times they should consider purchasing the book. "
                   "This can be an issue if you need all students to use the book in class or during open "
                   "book exams. Please reach out to the OER team if you have any questions about limited licenses.")]]
limit_ex_else = ("An Unlimited license means that an unlimited number of students can access "
                   "a book simultaneously.")
license_search_list = [['cop', 'limited user license --', True], ['nlimited', 'unlimited user license']]

# define functions
def create_dir(abs_pth:str, output_dir:str):
    '''
    abs_pth = the file path to the main directory scirpt is in, should be defined as global var
    output_dir = path from abs_path if there are intermediate directories and the output directory name
    creates a directory if none exists, if directory does exist, asks user for y/n input 
    if y, passes does not overwrite (but output files will be overwritten) 
    '''
    output_path = os.path.join(abs_pth, output_dir)
    if os.path.exists(output_path):
        x = input('output directory exists, files will be overwritten. Continue? y/n')
        if x == 'n':
            sys.exit('you chose to quit')
        else: 
            pass

    else: 
        os.mkdir(output_path)
def conditional_col(new_col_name:str, df, condition_col:str, search_text:list, else_text:str ):
    '''
    new_col_name = name of the new conditional col
    df = dataframe 
    condition_col = name of the existing column that new col is based on
    search text = list of lists, the first item in every sub list is the search text
        ie the text to search for in condition_col, the second item is the new col
        text, optional third item -- if there is a third item new col value will be
        search_text[1] + the value in the condition col (this was written for 
        license text). This can be a list containing only one list.
    else_text : if no search term is found, new_col will default to this value.
    '''
    df[new_col_name] = else_text
    for i in df.index:
        x = df.at[i, condition_col]
        for item in search_text:
            if x != else_text:
                if len(item)>2:
                    if item[0] in x:
                        df.at[i, new_col_name] = item[1] + x
                    else:
                        pass
                elif item[0] in x: 
                        df.at[i, new_col_name] = item[1]
            else:
                pass
            #else:
                #df.at[i, new_col_name] = else_text
def add_frq(df, col:str):
    '''
    df: dataframe you want to add a freq column to
    col: column to base freq count on -- will also be new col 
    name ie CRN_freq
    adds a count column based on anther column
    '''
    df[f'{col}_freq'] = df.groupby(col)[col].transform('count')
    return df           
def horizontal_explode(df, group1:str, group_2:list, cols:list, join_ch:str, output_file:str, flag_search, flag_else):
    '''
    df = df to group
    group1 = name of first col to group by, should be 'instructor_email_freq' 
    group2 = second cols to group by, could be a string, but expect a list of two cols
    cols = cols to include in output horizontal file, should be a list of col names 
    join_ch ; string/character to join on and then split on, ex: ";"
    output_file = the output file name in excel format, concated with group number ie {1}_book_emails.xlsx
    flag_search: used in conditional_col call for the search text, adding a limited license explaination
    flag_else: used in conditional col, adds the default unlimited license explaination
    '''
    # first form large groups of instr_freq -- each output file = 1 group
    df = df.groupby(group1)
    x_len = 0
    for name, group in df:
        x = group.groupby(group_2)[cols].agg(join_ch.join)
    # iterate over each col, and explode horizontally, add prefixes for exploded cols
        for col in x.columns:
            x = x.join(x[col].str.split(join_ch, expand=True).add_prefix(f'{col}_'))
       
        # print some stats
        print("group:",name, len(x.index),'total instructors')
        x_len = x_len + len(x.index)

        # add an explaination of the license type
        conditional_col('limit_explanation', x, 'license_text', flag_search, flag_else)

        # send to excel, file name based on group1
        x.to_excel(os.path.join(output_dir_path, f'{name}'+ output_file))

    print('Total unique instructors to email:',x_len)  
def save_deleted_rows(df_removed, df_dupes, cols:list, path, file_name:str):
    '''
    df_removed: df after filtering to remove duplicate rows (qc_matching)
    df_dupes: saved copy of df before deduping (qc_matching_dupes)
    cols: list of columns to filter to and merge on (need to lead to unique matches for merge)
    path: file path including abs path to save to
    file_name: file name to save to (deleted_dupes.xlsx)
    this function takes a slice of the deduped data, merges is back into un-deduped
    data with the indicator on, and then filters so you only have data that is ONLY
    in the un-deduped data ie a list of deleted rows.
    '''
    df_removed = df_removed.loc[:, cols]
    df_dupes = pd.merge(df_dupes, df_removed,
                        how='outer', on=cols, indicator=True )
    df_dupes = df_dupes.groupby('_merge').get_group('left_only')
    df_dupes.to_excel(os.path.join(path, file_name))
def save_cda_list(df, term_code:int, output:str, cols:list, check_col:str, ):
    '''
    df: dataframe
    term_code: integer version of current term
    output: output directory path
    cols: list of columns to include in cda list
    check_col: name of col to look for term_code in (ie 'Term_cda')
    this function always slices dataframe to grab rows that match 
    term_code and columns in cols list. Then dedupes on cols and saves 
    output as save_cda_term_code.xlsx
    '''
    x = df.loc[df[check_col] == term_code][cols]
    x = x.groupby(cols).first().reset_index()
    x.to_excel(os.path.join(output,f'save_cda_{term_code}.xlsx'))

# create output directory
create_dir(abs_pth, output_dir)

# add license text conditional col
conditional_col('license_text', qc_matching, 'DRM', license_search_list, 'limited user license')

# add course number column
qc_matching["Course_Number"] = qc_matching['Dept'] + " " + qc_matching['Sec'].astype(str)

len_before_dedupe = len(qc_matching) # get length for check later

# group by ISBN AND instructor email to add aggregate course number info
qc_matching = qc_matching.sort_values('Course_Number') # sort before so course number in correct order
qc_matching_dupes = qc_matching # save a copy w/ dupes for use later
isbn_instr_grouped = (qc_matching.groupby(["Instructor Email", "ISBN"])
        .agg({'Course_Number': lambda x: "/".join(x)})
        .rename({'Course_Number': 'course_numbers'}, axis=1).reset_index())

qc_matching = pd.merge(qc_matching, isbn_instr_grouped, how='left', on=["Instructor Email", "ISBN"])


# remove duplicate ISBN and Instructor rows (before adding freq columns)
# there definitely a better way to do this in just one step (combined above.
qc_matching = qc_matching.sort_values('Course_Number')
print("Before removeing duplicate ISBN & Instructor, df len is: ",len(qc_matching))
qc_matching = qc_matching.groupby(["Instructor Email", "ISBN"]).first().reset_index()
print("After removing dupes, df len is", len(qc_matching),'\n')

qc_matching = qc_matching.rename(columns={"Instructor Email": "instructor_email"}) # rename for ease 


# call add_freq to add CRN_freq and instructor_email_freq
qc_matching = add_frq(qc_matching, "CRN")
qc_matching = add_frq(qc_matching, "instructor_email")
# sort df
qc_matching = qc_matching.sort_values("instructor_email")
qc_matching = qc_matching.sort_values("CRN_freq")

# add instructor full name to get first-name last-name 
# (to split into two cols get rid of .map and add expand=True to split & add 2 cols)
qc_matching['instructor_full_name'] = qc_matching.Instructor.str.split(', ').map(lambda x : ' '.join(x[::-1]))

# add print link conditional col
qc_matching['print_libsearch_link'].fillna(" ", inplace=True) 
# first replace null values with spaces (for loop + groupby requires)
conditional_col('print_link?', qc_matching, 'print_libsearch_link', [['http', 'print link']], ' ')

# save a full copy of cleaned up and filtered list for convinient access to full list of emails
qc_matching.to_excel(os.path.join(output_dir_path,'deduped_titles_output.xlsx'))


# turn into mulitple horizontal files based on instructor
# slice full data to only get columns needed for sending the emails
t_grouped = qc_matching.loc[:, cols_to_keep]
t_grouped = t_grouped.applymap(str) # avoid join error by joining ints

# split by instructor_email_freq and then group by instructor email explode horiontally and save to excel
#horizontal_explode(t_grouped, 'instructor_email_freq', ['instructor_email', 'instructor_full_name']
                   #,email_cols, "; ", '_book_emails.xlsx',limit_ex_search, limit_ex_else )


# save a copy of de-duped / deleted rows
qc_matching_dupes = qc_matching_dupes.rename(columns={"Instructor Email": "instructor_email"})
#save_deleted_rows(qc_matching, qc_matching_dupes, 
                  #['instructor_email', 'Internal ID', 'CRN', 'ISBN']
                  #, output_dir_path, 'deleted_dupes.xlsx')
# save a clean list of the newly purchased and newly discovered titles to add to purchased_not_purchased
# identical to acquisitions list, but nicely pre-formatted. 
save_cda_list(qc_matching,term_int, output_dir_path, ['Title_y', 'ISBN', 'Purchased?', 
                'LibSearch Link', 'DRM', 'Term_cda', 'print_libsearch_link'], 'Term_cda')