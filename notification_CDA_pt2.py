import pandas as pd
import os
import sys

# define variables
term = "202401"
term_int = int(term)
abs_pth = os.path.dirname(os.path.abspath(__file__))
output_dir = f'notification_{term}/script2_output'
output_dir_path = os.path.join(abs_pth, output_dir)

# read data
qc_matching = pd.read_excel(os.path.join(abs_pth,f"notification_{term}/{term}_merge_matching_clean.xlsx"))
not_cleaned = pd.read_excel(os.path.join(abs_pth,f"notification_{term}/{term}_merge_matching.xlsx"))

# define column names (UPDATE HERE)
DRM = 'DRM' # DRM/license info column
dept = 'Dept' # department name or code ie the 'HIST' in HIST 101
CRN = 'CRN' # course section unique id
sec = 'Sec' # course section number ie the '101' in HIST 101
instructor_email = 'Instructor Email'
ISBN = 'ISBN'
instructor_name = 'Instructor' # instructor name (lastname, firstname)
ebook_link = 'LibSearch Link'
print_link = 'print_libsearch_link'  # print catalog link, if available
internal_id = 'Internal ID'  # unique identifier for books, can remove


# cols to keep for notifications: cols_to_keep is for both book list and email list, email_cols is for email only
cols_to_keep = ['instructor_email', 'Title_x', 'instructor_first_name', 'instructor_last_name',
                ebook_link,'print_link?',print_link, DRM,
                'Purchased?', 'course_numbers', 'license_text', 'CRN_freq', 
                'instructor_email_freq', CRN]
email_cols = ['Title_x', 'course_numbers', 'license_text', CRN]
save_cda_cols = ['Title_y', ISBN, 'Purchased?', 
                ebook_link, DRM, 'Term_cda', print_link]
save_cda_group = ['Title_y', ISBN,]

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
def concat_cols(new_col:str, df, col1:str, col2:str, join_ch:str):
    '''
    concatenates the values in two columns into one column as a string, seperated
    by a join character (make '' for no seperation), the original columns
    are left unchanged, and the new column is added to the dataframe. The values
    in both columns are converted to strings.

    new_col = name of new column
    df = dataframe
    col1 = name of first column to concatenate, must contain strings or ints
    col2 = name of second column to concatenate, must contain strings or ints
    join_ch = character to join on, ie ';' or ' '
    '''
    df[new_col] = df[col1].astype(str) + join_ch + df[col2].astype(str)
def add_frq(df, col:str):
    '''
    df: dataframe you want to add a freq column to
    col: column to base freq count on -- will also be new col 
    name ie CRN_freq
    adds a count column based on anther column
    '''
    df[f'{col}_freq'] = df.groupby(col)[col].transform('count')
    return df           
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
def save_cda_list(df, term_code:int, output:str, cols:list, group_cols:list, check_col:str):
    '''
    df: dataframe
    term_code: integer version of current term
    output: output directory path
    cols: list of columns to include in cda list
    group_cols: columns to group by (must contain unique values, no blanks)
    check_col: name of col to look for term_code in (ie 'Term_cda')
    this function always slices dataframe to grab rows that match 
    term_code and columns in cols list. Then dedupes on cols and saves 
    output as save_cda_term_code.xlsx
    '''
    x = df.loc[df[check_col] == term_code][cols]
    print(len(x))
    x = x.groupby(group_cols).first().reset_index()
    print("Save cda list (should equal acq + purchased): ",len(x))

    x.to_excel(os.path.join(output,f'save_cda_{term_code}.xlsx'))
# create output directory
create_dir(abs_pth, output_dir)

# add license text conditional col
conditional_col('license_text', qc_matching, DRM, license_search_list, 'limited user license')

# add course number column
concat_cols('Course_Number', qc_matching, 'Dept', 'Sec', ' ')
len_before_dedupe = len(qc_matching) # get length for check later

# aggregate course numbers for books used by the same instructor in mulitple courses (eg 4/500)
qc_matching = qc_matching.sort_values('Course_Number')  # sort before so course number in correct order
qc_matching = qc_matching.rename(columns={instructor_email: "instructor_email"}) # rename for ease 
qc_matching_dupes = qc_matching  # save a copy w/ dupes for use later
isbn_instr_grouped = (qc_matching.groupby(["instructor_email", "ISBN"])  # group by book and instrutor
        .agg({'Course_Number': lambda x: "/".join(x)})
        .rename({'Course_Number': 'course_numbers'}, axis=1).reset_index())
qc_matching = pd.merge(qc_matching, isbn_instr_grouped, how='left', on=["instructor_email", ISBN])


# remove duplicate rows on ISBN and Instructor (instructors using the same book in mulitple sections) 
qc_matching = qc_matching.sort_values('Course_Number')
print("Before removeing duplicate ISBN & Instructor, df len is: ",len(qc_matching))
qc_matching = qc_matching.groupby(["instructor_email", "ISBN"]).first().reset_index()
print("After removing dupes, df len is", len(qc_matching),'\n')


# add CRN_freq and instructor_email_freq
qc_matching = add_frq(qc_matching, CRN)
qc_matching = add_frq(qc_matching, "instructor_email")
# sort df
qc_matching = qc_matching.sort_values("instructor_email")
qc_matching = qc_matching.sort_values("CRN_freq")

# add instructor first-name last-name 
#qc_matching['instructor_full_name'] = qc_matching.Instructor.str.split(', ').map(lambda x : ' '.join(x[::-1])) # adds one full name col instead
qc_matching[['instructor_last_name', 'instructor_first_name']] = qc_matching[instructor_name].str.split(', ', expand=True) 

# add print link conditional col to include as linked text in emails
# rows without links must be blank
# replace null values with spaces (for loop + groupby requires)
qc_matching[print_link].fillna(" ", inplace=True) 
conditional_col('print_link?', qc_matching, print_link, [['http', 'print link']], ' ')

# save a full copy of cleaned up and filtered list for convinient access to full list of emails & books
qc_matching.to_excel(os.path.join(output_dir_path,'deduped_titles_full.xlsx'))

# save a clean list of the newly purchased and newly discovered titles to add to purchased_not_purchased
# identical to acquisitions list, but nicely pre-formatted. 
save_cda_list(not_cleaned,term_int, output_dir_path, save_cda_cols, save_cda_group, 'Term_cda')

# slice full data to only get columns needed for sending the emails
book_list = qc_matching.loc[:, cols_to_keep]
book_list = book_list.applymap(str) # avoid join error by joining ints

# save book list (1 of 2 files used for automated emails)
book_list.to_excel(os.path.join(output_dir_path,f'{term}_books.xlsx'))

# create email list 
# group by instructor to remove duplicates, aggregate other columns
join_ch= ';'
email_list = book_list.groupby(['instructor_email','instructor_first_name', 'instructor_last_name', 
                                'instructor_email_freq'])[email_cols].agg(join_ch.join)
# add a license explanation column based on license text aggregation
conditional_col('limit_explanation', email_list, 'license_text', limit_ex_search, limit_ex_else)
# save email list (2 of 2 files used for automated emails)
email_list.to_excel(os.path.join(output_dir_path,f'{term}_emails.xlsx'))


# save a copy of de-duped / deleted rows
save_deleted_rows(qc_matching, qc_matching_dupes, 
                  ['instructor_email', internal_id, CRN, ISBN]
                  , output_dir_path, 'deleted_dupes.xlsx')