import os
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# read in data
sb_df = pd.read_excel("/Users/aservice/Documents/python/faculty_publications_SB.xlsx")
ds_df = pd.read_excel("/Users/aservice/Documents/python/ds_faculty_list.xlsx")

mat1 = []
mat2 = []

threshold = 90 # set threshold for fuzzy matching, update here
# remove null values
sb_df = sb_df.dropna(subset=['instructor'])
ds_df = ds_df.dropna(subset=['instructor'])

sb_df['instructor'] = sb_df['instructor'].str.lower()
ds_df['instructor'] = ds_df['instructor'].str.lower()
# convert dfs to lists

sb_list = sb_df['instructor'].tolist()
ds_list = ds_df['instructor'].tolist()

#print(sb_list, ds_list)

# iterate through sb_list to extract closest match from ds_list

for i in sb_list: 
    mat1.append(process.extract(i, ds_list, limit=2))

sb_df['matches'] = mat1

# iterate through the closest matches and filter out to get the best match
p = []
for j in sb_df['matches']:
    for k in j:
        # print(k)
        # print(k[1])
        if k[1] >= threshold:
            p.append(k[0])
    mat2.append(";".join(p))
    p = []
# store matches in sb_df
sb_df['matches'] = mat2

#print(sb_df.head(35))
print("length of matches",sb_df.count())
print("length of insturctors",len(sb_df['instructor']))

sb_df.to_excel("/Users/aservice/Documents/python/fuzzy_export2.xlsx")