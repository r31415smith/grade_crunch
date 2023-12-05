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
 -- i could offer to accept professor's gradebook and then interate the labs into their gradebook. I pretty much already have the code for that from BB stuff...
 

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


# load all the the xlsx files in a folder
# dir_name = r'G:\My Drive\@courses\@Intro labs - manage\@lab scores\PHYS 135 expt'
dir_name = './fa_23 lab scores/PHYS 135 lab scores/'
data_folder = Path(dir_name)
filename_suffix = 'csv'

listfiles = sorted(os.listdir(data_folder) )
listfiles1 = [elem for elem in listfiles if elem[-3:] == filename_suffix] 

def load_data(filename):
    file_to_open = data_folder / filename
    # skip the second line because this is irrelevant header info:
    section_data = pd.read_csv(file_to_open, delimiter=',',encoding = "ISO-8859-1",skiprows = [1]) 
    section_data['class avg'] = section_data['Final Points'].mean()
    section_data = section_data.rename(columns={'SIS Login ID': 'NetID'})
    return section_data

raw_data_sets = [load_data(filename) for k,filename in enumerate(listfiles1)]







# Function to calculate total excluding the lowest score and treating NaN as 0
def sum_without_lowest(scores):
    scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    if len(scores_no_nan) > 1:
        return sum(sorted(scores_no_nan, reverse=False)[1:])  # Sum all scores except the lowest
    else:
        return sum(scores_no_nan)  # If only one score or no scores, return that score
    
# Function to calculate total excluding the lowest score and treating NaN as 0
def sum_all(scores):
    scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    return sum(scores_no_nan)  # If only one score or no scores, return that score

def determine_num_missed(scores):
    scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    # Increment scores less than 5.1 by 1 (because some book-keeping marks a 5 as no-show)
    missed_which = [1 if x < 5.1 else 0 for x in scores_no_nan]
    num_missed = sum(missed_which)
    return num_missed

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
    # Calculate the number of labs missed:
    for index, row in sect_num.filter(regex='NB', axis=1).iterrows():
        list_missed.append(determine_num_missed(row))
    sect['# labs missed'] = list_missed
    
    # add up prelabs
    total_scores = []
    for index, row in sect_num.filter(regex='prelab', axis=1).iterrows():
        total_scores.append(sum(row))
    sect['prel total'] = total_scores   
    
    # add up extra credit
    total_scores = []
    for index, row in sect_num.filter(regex='EC', axis=1).iterrows():
        total_scores.append(sum(row))
    sect['Extra credit total'] = total_scores
   
    sect['total score'] = sect['total notebook lowest dropped'] + sect['Extra credit total'] + sect['prel total']
    
    sect['total %'] = sect['total score'] / (sect_num.filter(regex='NB', axis=1).shape[1]*10 +  sect_num.filter(regex='prelab', axis=1).shape[1]*2)*100
    
    section_avg = np.mean(sect['total %'] )
    section_std = np.std(sect['total %'] )
    sect['z-score'] =  ( sect['total %']-section_avg ) / section_std
    sect['section avg'] = section_avg
    section_processed = sect[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', 'class avg','# labs missed','z-score','section avg']] 
    
    return section_processed


proc_data_sets = [gather_scores(sect) for sect in raw_data_sets]


#########3 concat all the data from all sections:
data_concat = pd.concat(proc_data_sets,sort=True,ignore_index=True)
# sort the contatenated result alphabetically
result = data_concat[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', 'section avg','# labs missed','z-score']] 
result = result.sort_values(by=['Student'])
result = result[~result['Student'].str.contains('Student, Test', regex=True)]








############# determine a curve so that each section has the same average:
    
# determine the curve factor for each student
# (not super efficient but doesn't matter here)
highest_avg = result['section avg'].max()

for index,row in result.iterrows():
    x = (100 - highest_avg) / (100 - result['section avg'])
    result['x_val'] = x
    result['Total curved grade %'] =  result['total %']*x + 100*(1-x)

# testing: get the "new" average for each section...  (not so easy because sections are already all mixed in...)

# make a result that is exactly the cols that we want
trimmed_result = result[['Student','NetID', 'total %', '# labs missed', 'z-score',
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


filename_to_save = 'Compiled_grades_'+data_folder.parts[-1][:8]+'.xlsx'
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
    

    # # Apply conditional formatting to highlight cells where 'flag' > 4
    # worksheet.conditional_format('B2:ZZ1048576', {'type': 'cell',
    #                                                'criteria': '>',
    #                                                'value': 4,
    #                                                'format': red_format,
    #                                                'strict': True,
    #                                                'stop_if_true': False})
# # Close the Pandas Excel writer and output the Excel file happens when we leave the "with" scope







# path_to_save = data_folder
# filename_to_save = 'Result2.xlsx'
# file_to_save = path_to_save / filename_to_save
# # Create a Pandas Excel writer using XlsxWriter as the engine.
# with pd.ExcelWriter(file_to_save, engine="xlsxwriter") as writer:
#     result.to_excel(writer, index=False,sheet_name='Sheet1')
#     # worksheet = writer.sheets['Sheet1']




