#import pandas module
import pandas as pd
#read the csv file into a dataframe


lib_spending = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/fy24_25_cda_spending_raw.xlsx")
all_cda = pd.read_excel("/Users/aservice/Library/CloudStorage/OneDrive-UniversityOfOregon/CDA/lib_spending/all_cda_202301-202404.xlsx")



#define variables
expld_on = "ISBN" #column name to explode on
#delimiter = ";"
delimiter=";" #new line char \n

#split the column to explode into a list and add that column to the data frame

lib_spending[expld_on] = lib_spending[expld_on].str.split(delimiter, expand=False)

#print(df_p.head()['ISBN'])

df_main = lib_spending.explode(expld_on)

#print(df_main.keys(), df_main.shape)
#print(df_main['ISBN'].head())
df_main['ISBN'] = df_main['ISBN'].astype(str)
#print(all_cda['ISBN'].head())   
#print(df_main['ISBN'].head())
all_cda['ISBN'] = pd.to_numeric(all_cda['ISBN'], errors='coerce').astype('Int64')
all_cda['ISBN'] = all_cda['ISBN'].astype(str)
#print(all_cda['ISBN'].head())   
df_main['ISBN'] = df_main['ISBN'].str.strip()


df_main = df_main.merge(all_cda,
                     on = 'ISBN',
                     how = 'outer',
                     indicator=True)

#print(df_main)
#print(df_main.keys())
df_main.sort_values(by=['Title_y'], inplace=True)
#df_main.to_excel("exploded_pre_drop.xlsx")
cda_right = df_main.loc[df_main['_merge'] == 'right_only'].copy() 

#print('GOT HERE OK')
#print(cda_right.keys())
cda_right['Title_y'] = cda_right['Title_y'].apply(lambda x:x.lower() if isinstance(x, str) else x)
cda_right['Title_y'] = cda_right['Title_y'].str.replace(' : ', ' ')
cda_right['Title_y'] = cda_right['Title_y'].str.replace(':','')
cda_right['Title_y'] = cda_right['Title_y'].str.replace("'", "")
cda_right['Title_y'] = cda_right['Title_y'].str.replace('-', ' ')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('+', ' ')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('/', '')
cda_right['Title_y'] = cda_right['Title_y'].str.replace(',','')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('!', '')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('?', '')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('.', '')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('"', '')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('*', '')
cda_right['Title_y'] = cda_right['Title_y'].str.replace('%', '')
#cda_right['Title_y'] = cda_right['Title_y'].str.strip('the ') # doesn't work, removes the 's at the end of words
cda_right['Title_y'] = cda_right['Title_y'].str.replace(r'^the ', '', regex=True)

df_main.drop_duplicates(subset=['PO Line Reference'], keep='first', inplace=True) # also drops all but one right onlys because the POLs are blanks

both = df_main[df_main['_merge'] == 'both'].copy()
spending_left = df_main[df_main['_merge'] == 'left_only'].copy()
spending_left.drop(columns=['_merge'], inplace=True)
cda_right.drop(columns=['_merge'], inplace=True)

title_merge = spending_left.merge(cda_right,
                     left_on = 'Title_x', right_on= 'Title_y',
                     how = 'outer',
                     indicator=True)

title_merge_both = title_merge[title_merge['_merge'] == 'both']
unmatched_left = title_merge[title_merge['_merge'] == 'left_only']
unmatched_right = title_merge[title_merge['_merge'] == 'right_only']
print(unmatched_right.keys())
unmatched_right = unmatched_right[['ISBN_y', 'Title_y_y', '_merge']]

df_main = df_main[df_main['_merge'] == 'both'].copy() # keep only matched rows so when we merge in additional title matches there aren't dupes
df_main = pd.concat([df_main, title_merge_both]) 
# all the unmatched rows are graped first into the spending_l/r and then into the unmatched_left/right to review
#title_merge.to_excel('title_merge3.xlsx')
unmatched_left.to_excel('unmatched_left.xlsx')
unmatched_right.to_excel('unmatched_right.xlsx')

#print(df_main.head())
df_main.to_excel("exploded_isbn-title.xlsx")