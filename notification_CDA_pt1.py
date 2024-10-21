

# import pandas module
import pandas as pd
import os

# define variable, CHANGE TERM
term = "202401" # term is used in file paths throughout

# get absolute path
abs_pth = os.path.dirname(os.path.abspath(__file__))

# read in data
acq_own = pd.read_excel(os.path.join(abs_pth, f"notification_{term}/{term}_CDA_acq_own.xlsx"))
acq_purchased = pd.read_excel(os.path.join(abs_pth, f"notification_{term}/{term}_CDA_acq_purchased.xlsx"))
acq_past = pd.read_excel(os.path.join(abs_pth, "all_titles_purchased_not_purchased.xlsx"))
bookstore = pd.read_excel(os.path.join(abs_pth, f"notification_{term}/{term}_full_ds.xlsx"))
print("DS data length: ",len(ds))

#print(bookstore.keys(), acq_own.keys(), acq_purchased.keys(), acq_past.keys()) # prints all column names for error checking

# add purchased column based on file name
acq_own["Purchased?"] = "owned/access"
acq_purchased["Purchased?"] = "purchased"

# clean past purchased CDA
acq_past= acq_past.rename(columns={"Title_cda" : "Title"})
acq_past = acq_past.loc[(acq_past["Purchased?"]!= "no")]

# Append acquisitions data into one dataframe
acq = pd.concat([acq_own, acq_purchased, acq_past])
print("All acq length (unique books owned or purchased): ", len(acq))

# fill Term_cda
acq["Term_cda"].fillna(term, inplace=True)

# Merge Data from Book Store with Acquisition Data
print("DS data length: ",len(bookstore))

# remove no eBook Allowed titles
bookstore = bookstore[bookstore["No eBook Allowed"] !="Yes"]

# merge Book Store data and acqusitions data
matching = bookstore.merge(acq,
                     on = "ISBN",
                     how = 'outer')
print("all merged titles length: ", len(matching))
                    
matching = matching[~matching["DRM"].isna()] #select only rows in column 'DRM' with values-- 
# only rows with matching acq record
matching = matching.sort_values("DRM")

# add lowercase title columns for QC
matching['Title_cda'] = matching['Title_y'].str.lower()
matching['Title_ds'] = matching['Title_x'].str.lower()
matching['Title_ds'] = matching["Title_ds"].fillna("no DS record")


def flag_dif_col(df, col_1, col_2, new_col):
    '''
    df: dataframe
    col_1: first col to compare
    col_2: second col to compare
    new_col: new col name to add
    updates df in place, adds a new flag col which tells you
    the number of words that are different in col_1 and col_2
    '''
    for i in df.index:
        x = df.at[i, col_1]
        y = df.at[i, col_2]
        z = 0
        x = x.split()
        y = y.split()
        for word in range(min(len(x), len(y))):
            if x[word] == y[word]:
                pass
            else:
                z = z + 1
        df.at[i, new_col] = z
# add title_flag, # of different words in cda vs ds titles
flag_dif_col(matching, 'Title_cda', 'Title_ds', 'title_flag')       

# save for QC and use in next script
matching.to_excel(os.path.join(abs_pth, f"notification_{term}/{term}_merge_matching.xlsx"))
print("All matching length: ", len(matching))

# print some stats -- could save them instead

owned_access_total = len(acq_own)
purchase_total = len(acq_purchased)
print('total books owned/accessed: ', owned_access_total)
print('Total books purchased: ', purchase_total)
print("Complete QC1 before saving file and running notification script 2", 
      f"\nFile to QC is saved as: notification_{term}/{term}_merge_matching.xlsx")

