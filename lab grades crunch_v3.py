# -*- coding: utf-8 -*-
"""
Script for crunching all lab scores (all sections CSV in a single folder) and computing a "fair curve" that brings each section up to the highest average.
Instructor can then decide how they want to use this info.

the improvement here was to do formatting (col width). 
future: take in the gradebook from lecturer and append these...
could use from:
    "@Merge labs to lecture BB-simple"
in grading - Remark   - https://jupyterhub.csueastbay.edu/user/ex9932/notebooks/grading%20-%20Remark/%40Merge%20labs%20to%20lecture%20BB-simple.ipynb#

input BB gradebook CSV and lab master CSV, and it outputs a grade file similar to the BB gradebook, can then be easily copied over...

assumptions:
* CSV files are coming from "export entire gradebook"
* the names of columns are not changed from what I initiated in Canvas
* instructors can either fill in values or leave blank, it doesn't matter. -- but in order to count the missing labs, those missing labs need to have a score of 5 or less.
* folder name starts with the name of the class -- e.g. "PHYS 125 lab scores"
* Extra credit sections are ok. If extra credit is awarded, these columns can be labeled "EC 1" etc., as long as they have the expression "EC" in them it's fine -- or extra credit could be applied to a particular lab or prelab




* If you want to do the calculations on your own and report the total scores, .... (ugh not so easy, needs more programming)


* future: if there are multiple lecture sections, instructors need to go through the list
 -- i could input lecture section instructor's gradebook and then interate the labs into their gradebook. I pretty much already have the code for that from BB stuff...
 

"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import norm
import os.path
from pathlib import Path
import xlrd
import openpyxl
import xlsxwriter


# load all the the CSV files in a folder
dir_name = './fa_23 lab scores/PHYS 126 lab scores/'
data_folder = Path(dir_name)
filename_suffix = 'csv'

listfiles = sorted(os.listdir(data_folder) )
# only include files with the CSV suffix
listfiles1 = [elem for elem in listfiles if elem[-3:] == filename_suffix] 

def load_data(filename):
    file_to_open = data_folder / filename
    # skip the second line because this is irrelevant header info:
    section_data = pd.read_csv(file_to_open, delimiter=',',encoding = "ISO-8859-1",skiprows = [1]) 
    section_data['class avg'] = section_data['Final Points'].mean()
    section_data = section_data.rename(columns={'SIS Login ID': 'NetID'})
    return section_data

# make a list of dataframes for each section scores (but still keep distinguished from each other)
raw_data_sets = [load_data(filename) for k,filename in enumerate(listfiles1)]




# Example function to calculate total excluding lowest score
def sum_without_lowest(row):
    # Find the indices of sorted values
    sorted_indices = np.argsort(row)
    # print(row) 
    # print(row.index[sorted_indices])
    # Exclude the lowest score unless it corresponds to 'practical'
    if " practical" in row.index[sorted_indices[0]]:
        lowest_score = row[sorted_indices[0]]
        second_lowest_score = row[sorted_indices[1]]
        total = row.sum() - second_lowest_score
    else:
        lowest_score = row[sorted_indices[0]]
        total = row.sum() - lowest_score
    return total


# # Function to calculate total excluding the lowest score and treating NaN as 0
# def sum_without_lowest(scores):
#     # print(scores)
#     # makes sure to only drop lowest lab where 'NB' shows up, but will not drop the lowest lab if 'practical' does shows up there
#     # i.e., lowest lab as practical will not be dropped.
#     scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
#     if len(scores_no_nan) > 1:
#         scores_sorted = sorted(scores_no_nan)
#         # Check if the last column contains the string "practical" (case insensitive)
#         last_column_name = scores_sorted[-1]
#         last_contains_practical = last_column_name.str.contains('practical', case=False).any()
#         if last_contains_practical:
#             sum_except_second_to_last = scores_sorted.iloc[:].sum()-scores_sorted.iloc[-2]
#             return sum_except_second_to_last
#         else:
#             return sum(sorted(scores_no_nan, reverse=False)[1:])  # Sum all scores except the lowest
#     else:
#         return sum(scores_no_nan)  # If only one score or no scores, return that score

# # Function to calculate total excluding the lowest score and treating NaN as 0
# def sum_without_lowest(scores):
#     # print(scores)
#     # makes sure to only drop lowest lab where 'NB' shows up, but will not drop the lowest lab if 'practical' does shows up there
#     # i.e., lowest lab as practical will not be dropped.
#     scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
#     if len(scores_no_nan) > 1:
#         return sum(sorted(scores_no_nan, reverse=False)[1:])  # Sum all scores except the lowest
#     else:
#         return sum(scores_no_nan)  # If only one score or no scores, return that score
    
    
    
# Function to calculate total excluding the lowest score and treating NaN as 0
def sum_all(scores):
    scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    return sum(scores_no_nan)  # If only one score or no scores, return that score

# counts how many labs were missed by each student
def determine_num_missed(scores):
    scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    # Increment scores less than 5.1 by 1 (because some book-keeping marks a 5 as no-show)
    missed_which = [1 if x < 5.1 else 0 for x in scores_no_nan]
    num_missed = sum(missed_which)
    return num_missed

# Function to count "valid" columns
def count_valid_columns(df, threshold):
    below_threshold = []
    for col in df.columns:
        below_threshold.append((df[col] < threshold).sum() > (len(df[col]) / 2))
    return len(df.columns) - sum(below_threshold)


######## crunch those numbers! get scores for each section (pre-curve)
def gather_scores(sect):
    # Convert non-numeric values to NaN
    sect_num = sect.apply(pd.to_numeric, errors='coerce')
    
    # Replace NaN values with 0
    sect_num = sect_num.fillna(0)
    
    
    # Calculate total excluding the lowest score for each student and add to 'total' column
    total_scores = []
    for index, row in sect_num.filter(regex='NB', axis=1).iterrows():
        total_scores.append(sum_all(row))
    sect['total notebook score'] = total_scores
    
    # Calculate total excluding the lowest score for each student and add to 'total' column
    total_scores = []
    for index, row in sect_num.filter(regex='NB', axis=1).iterrows():
        total_scores.append(sum_without_lowest(row))
    sect['total notebook lowest dropped'] = total_scores

    list_missed = []
    # Calculate the number of labs missed for each student:
    for index, row in sect_num.filter(regex='NB', axis=1).iterrows():
        list_missed.append(determine_num_missed(row))
    sect['# labs missed'] = list_missed
    
    # add up prelab scores for each student:
    total_scores = []
    for index, row in sect_num.filter(regex=r'^(?=.*[0-9])[pP]relab', axis=1).iterrows():  #filter(regex='[pP]relab', axis=1).iterrows():
        total_scores.append(sum(row))
    sect['prel total'] = total_scores   
    

    # Count the number of "valid" columns -- meaning more than half the students got more than half the credit
    num_valid_columns_NB = count_valid_columns(sect_num.filter(regex='NB', axis=1), threshold=4.9)
    # include that a number needs to appear in the title of column for prelabs...
    num_valid_columns_prelab = count_valid_columns(sect_num.filter(regex=r'^(?=.*[0-9])[pP]relab', axis=1), threshold=1)
   
    print(num_valid_columns_NB)
    print(num_valid_columns_prelab)
    # add up extra credit for each student:
    total_scores = []
    for index, row in sect_num.filter(regex='EC', axis=1).iterrows():
        total_scores.append(sum(row))
    sect['Extra credit total'] = total_scores
   
    sect['total score'] = sect['total notebook lowest dropped'] + sect['Extra credit total'] + sect['prel total']
    
    # calculate score as a %, after dropping lowest, and considering N-1 "valid" lab NB and prelab scores.
    sect['total %'] = sect['total score'] / ((num_valid_columns_NB-1)*10 +  (num_valid_columns_prelab-1)*2)*100
    
    section_avg = np.mean(sect['total %'] )
    section_std = np.std(sect['total %'] )
    sect['z-score'] =  ( sect['total %']-section_avg ) / section_std
    sect['section avg'] = section_avg
    section_processed = sect[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', 'class avg','# labs missed','z-score','section avg']] 
    
    return section_processed


# make "processed" data frames that now include summary information
proc_data_sets = [gather_scores(sect) for sect in raw_data_sets]


######### concatenate all the data from all sections:
data_concat = pd.concat(proc_data_sets,sort=True,ignore_index=True)

result = data_concat[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', 'section avg','# labs missed','z-score']] 
# remove the "test student"
result = result[~result['Student'].str.contains('Student, Test', regex=True)]
# sort the contatenated result alphabetically
result = result.sort_values(by=['Student'])








############# determine a curve so that each section has the same average:
    
# determine the curve factor for each student
# (not super efficient but doesn't matter here)
highest_avg = result['section avg'].max()

for index,row in result.iterrows():
    x = (100 - highest_avg) / (100 - result['section avg'])
    result['x_val'] = x
    result['Total curved grade %'] =  np.minimum(result['total %']*x + 100*(1-x),100)

# testing: get the "new" average for each section...  (not so easy because sections are already all mixed in...)

# make a result that is exactly the cols that we want
trimmed_result = result[['Student','NetID', 'total %', '# labs missed', 'z-score',
      'Total curved grade %']].sort_values(by=['Student'])  # 'z-score', 'class avg',

longer_result = result[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', '# labs missed', 'z-score',
      'Total curved grade %']].sort_values(by=['Student'])  # 'z-score', 'class avg',


############# save the trimmed summary file as Excel worksheet:

# save the concatenated excel files into one big file in the same folder:
path_to_save = data_folder
# save the concatenated excel files into one big file in a subfolder called "results"
path_to_save = data_folder # / 'result' 
# Check whether the specified path exists or not
isExist = os.path.exists(path_to_save)
if not isExist:
  # Create a new directory because it does not exist 
  os.makedirs(path_to_save)


filename_to_save = 'Compiled_grades2_'+data_folder.parts[-1][:8]+'.xlsx'
file_to_save = path_to_save / filename_to_save


# Create a Pandas Excel writer using XlsxWriter as the engine.
with pd.ExcelWriter(file_to_save, engine="xlsxwriter") as writer:
    trimmed_result.to_excel(writer, index=False,sheet_name='Sheet1',float_format = '%.2f')
    worksheet = writer.sheets['Sheet1']
    
    # Set the column width
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 4, 10)
    # worksheet.set_column(3, 3, 10)
    # worksheet.set_column(4,4, 7)
    worksheet.set_column(45,5, 16)
    
    workbook = writer.book
    # Add a format with red fill for values meeting the condition (flag > 4)
    red_format = workbook.add_format({'bg_color': '#FF9A98'})


    # Apply conditional formatting to column D based on the condition (greater than 4)
    worksheet.conditional_format('D2:D' + str(trimmed_result.shape[0]), {'type': 'cell',
                                                           'criteria': '>',
                                                           'value': 4,
                                                           'format': red_format})
    longer_result.to_excel(writer, index=False,sheet_name='Sheet2',float_format = '%.2f')
    worksheet = writer.sheets['Sheet2']
    
 




# path_to_save = data_folder
# filename_to_save = 'Result2.xlsx'
# file_to_save = path_to_save / filename_to_save
# # Create a Pandas Excel writer using XlsxWriter as the engine.
# with pd.ExcelWriter(file_to_save, engine="xlsxwriter") as writer:
#     result.to_excel(writer, index=False,sheet_name='Sheet1')
#     # worksheet = writer.sheets['Sheet1']




