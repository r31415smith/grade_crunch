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


* If you want to do the calculations on your own and report the total scores, maybe possible to integrate.... (but not so easy, needs more programming)


* future: if there are multiple lecture sections, instructors need to go through the list
 -- i could input lecture section instructor's gradebook and then interate the labs into their gradebook. I pretty much already have the code for that from BB stuff that I could convert...
 

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
# dir_name = './fa_23 lab scores/PHYS 125 lab scores/'
# dir_name = './fa_23 lab scores/PHYS 125 lab scores SS/'
# dir_name = './fa_23 lab scores/PHYS 126 lab scores/'
dir_name = './fa_23 lab scores/PHYS 135 lab scores/'
# dir_name = './fa_23 lab scores/PHYS 136 lab scores/'

data_folder = Path(dir_name)
filename_suffix = 'csv'

listfiles = sorted(os.listdir(data_folder) )
# only include files with the CSV suffix
listfiles1 = [elem for elem in listfiles if elem[-3:] == filename_suffix] 

def load_data(filename):
    file_to_open = data_folder / filename
    # skip the second line because this is irrelevant header info:
    section_data = pd.read_csv(file_to_open, delimiter=',',encoding = "ISO-8859-1",skiprows = [1]) 
    section_data['raw avg'] = section_data['Final Points'].mean()
    section_data = section_data.rename(columns={'SIS Login ID': 'NetID'})
    return section_data

# make a list of dataframes for each section scores (but still keep distinguished from each other)
raw_data_sets = [load_data(filename) for k,filename in enumerate(listfiles1)]




# Function to calculate total excluding lowest score
def sum_without_lowest(row):
    # Find the indices of sorted values
    sorted_indices = np.argsort(row)
    # print(row) 
    # print(row[sorted_indices])
    # Exclude the lowest score unless it corresponds to 'practical'
    if "practical" in row.index[sorted_indices[0]]:
        lowest_score = row[sorted_indices[0]]
        second_lowest_score = row[sorted_indices[1]]
        total = row.sum() - second_lowest_score
    else:
        lowest_score = row[sorted_indices[0]]
        total = row.sum() - lowest_score
    # print(total)
    return total


    
# Function to calculate total excluding the lowest score and treating NaN as 0
def sum_all(scores):
    scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    return sum(scores_no_nan)  # If only one score or no scores, return that score

# counts how many labs were missed by each student
def determine_num_missed(scores):
    scores_no_nan = scores   # just keep all the scores, even if they are 0...
    # scores_no_nan = [score for score in scores if score != 0]  # Exclude 0 (NaN values after conversion)
    # Increment scores less than 5.1 by 1 (because some book-keeping marks a 5 as no-show)
    missed_which = [1 if x < 5.1 else 0 for x in scores_no_nan]
    num_missed = sum(missed_which)
    # print(scores_no_nan)
    # print(missed_which)
    return num_missed

# a "valid" column is where more than half the entries are greater than the threshold value.
# Function to get "valid" columns
def get_valid_columns(df, threshold):
    valid_columns = []
    for col in df.columns:
        if (df[col] > threshold).sum() > (len(df[col]) / 2):
            valid_columns.append(col)
    return df[valid_columns]  # Return a DataFrame containing only the valid columns




######## crunch those numbers! get scores for each section (pre-curve)
def gather_scores(sect):
    # Convert non-numeric values to NaN
    sect_num = sect.apply(pd.to_numeric, errors='coerce')
    
    # Replace NaN values with 0
    sect_num = sect_num.fillna(0)
    
    
    
    ############ Lab Notebooks (submissions) ############
    # filter out lab NB (or 'submission') cols:
    # include that a number needs to appear in the title of column for prelabs...
    labNB_cols = sect_num.filter(regex=r'(?i)(submission|NB)', axis=1)
    # print(labNB_cols)
    # print(labNB_cols.shape[1])
    # column_headers = labNB_cols.columns.values.tolist()
    # print("The Column Header :", column_headers)
    # validate: all prelab columns should have more than half the students getting a score of more than 5 (half the credit), or it doesn't count.
    valid_columns_lab_NB = get_valid_columns(labNB_cols, threshold=5)
    
    num_valid_columns_NB = valid_columns_lab_NB.shape[1]
    # print(num_valid_columns_NB)
    
    # Calculate total of all labs for each student
    # Calculate total excluding the lowest score for each student
    # Calculate the number of labs missed for each student
    total_scores = []
    total_score_lowest_dropped = []
    list_missed = []
    for index, row in valid_columns_lab_NB.iterrows():
        total_scores.append(sum_all(row))
        total_score_lowest_dropped.append(sum_without_lowest(row))
        list_missed.append(determine_num_missed(row))
    sect['total notebook score'] = total_scores
    sect['total notebook lowest dropped'] = total_score_lowest_dropped
    sect['# labs missed'] = list_missed
    

    
    ############      prelabs       ############
    # filter out potential pelab cols:
    # include that a number needs to appear in the title of column for prelabs...
    prelab_cols = sect_num.filter(regex=r'(?i)^(?=.*\d)(?=.*prelab)', axis=1)
    
    # validate: all prelab columns should have more than half the students getting a score of more than 1, or it doesn't count.
    valid_columns_prelab = get_valid_columns(prelab_cols, threshold=1)
    num_valid_columns_prelab = valid_columns_prelab.shape[1]
    
    # add up valid prelab scores for each student:
    # print(valid_columns_prelab)
    total_scores = []
    for index, row in valid_columns_prelab.iterrows(): 
        total_scores.append(sum(row))
    sect['prel total'] = total_scores   
    

   
    ############ summarize all the data for a given section  ############
    print('# labs:    ',num_valid_columns_NB)
    print('# prelabs: ',num_valid_columns_prelab)
    # add up extra credit for each student:
    total_scores = []
    for index, row in sect_num.filter(regex='EC', axis=1).iterrows():
        total_scores.append(sum(row))
    sect['Extra credit total'] = total_scores
   
    sect['total score'] = sect['total notebook lowest dropped'] + sect['Extra credit total'] + sect['prel total']
    sect['lab NB %'] = sect['total notebook lowest dropped']/((num_valid_columns_NB-1)*10)*100
    sect['prelab %'] = sect['prel total']/((num_valid_columns_prelab)*2)*100
    # calculate score as a %, after dropping lowest, and considering N-1 "valid" lab NB and prelab scores.
    sect['total %'] = sect['total score'] / ((num_valid_columns_NB-1)*10 +  (num_valid_columns_prelab-1)*2)*100
    
    section_avg = np.mean(sect['total %'] )
    section_std = np.std(sect['total %'] )
    sect['z-score'] =  ( sect['total %']-section_avg ) / section_std
    sect['section avg'] = section_avg
    sect['section std']=section_std
    section_processed = sect[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', 'raw avg','# labs missed','z-score','section avg','section std','lab NB %','prelab %']]
    
    return section_processed





# make "processed" data frames that now include summary information
proc_data_sets = [gather_scores(sect) for sect in raw_data_sets]

# this is a record for me internally to see what the section averages were...
sect_avgs = [proc_data_sets[index]['section avg'][0] for index,sect in enumerate(raw_data_sets)]

######### concatenate all the data from all sections:
data_concat = pd.concat(proc_data_sets,sort=True,ignore_index=True)

result = data_concat  #[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','Extra credit total','total %', 'section avg','# labs missed','z-score']] 
# remove the "test student"
result = result[~result['Student'].str.contains('Student, Test', regex=True)]
# sort the contatenated result alphabetically
result = result.sort_values(by=['Student'])






############# determine a curve based on z-scores so that each section has the same average:

# determine the scaling based on z-scores for each student
# (not super efficient but doesn't matter here)
highest_avg = result['section avg'].max()

# Adjust grades to achieve a common desired average score
desired_average = highest_avg  # Set desired average score

# Scale Z-scores to achieve the desired average score
scaled_z_score = result['z-score'] * (desired_average / result['section avg'])
result['scaled_z_score'] = scaled_z_score
# Convert scaled Z-scores back to original scores
adjusted_score = scaled_z_score * result['section std'] + result['section avg']

# Now 'adjusted_scores' contains the adjusted grades for each lab section
# added rule: make sure that no one's grade is actually lowered using this method. 
# we will take the adjusted score, or their original, whichever is higher:
result['Total curved grade %'] = np.maximum(adjusted_score,result['total %'])

# flatten grade if over 100
result['Total curved grade %'] = np.minimum(np.maximum(adjusted_score,result['total %']),100)
# result['Total curved grade %'] = adjusted_score





# make a result that is exactly the cols that we want
trimmed_result = result[['Student','NetID', 'total %', '# labs missed', 'z-score',
      'Total curved grade %']].sort_values(by=['Student'])  # 'z-score', 'class avg',

longer_result = result[['Student','NetID', 'total notebook score','total notebook lowest dropped','prel total','lab NB %','prelab %','Extra credit total','total %', '# labs missed', 'z-score',
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
    trimmed_result.to_excel(writer, index=False,sheet_name='Lab summary',float_format = '%.2f')
    worksheet = writer.sheets['Lab summary']
    
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
    longer_result.to_excel(writer, index=False,sheet_name='Lab more detailed',float_format = '%.2f')
    worksheet = writer.sheets['Lab more detailed']
    
 
