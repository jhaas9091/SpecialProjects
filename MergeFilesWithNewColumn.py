"""Merge all xlsx files into one master spreadsheet
Insert origin source (file name) as first column"""

import os
import pandas as pd

# copy&paste directory address of source folder
path = r'C:\Users\jbae1\Downloads\Dummy Files'
# copy&paste directory address of where you want to save output file
outputfile = r'C:/Users/jbae1/OneDrive/Desktop/Output.xlsx'

# an empty data frame for new output excel file
mergedxlsx = pd.DataFrame()
mergedxlsx['File Name'] = 'File Name' # appends new column with file name

# limits down to files with excel extension
for root, subdir, file in os.walk(path):
    for file in file:
        if file.endswith('.xlsx'):

            df = pd.read_excel((os.path.join(root, file)), header=0, dtype=str)  # reads excel file
            workbook = pd.DataFrame(df)  # assigns data frame to read excel file
            workbook = workbook.set_axis(['timestamp', 'ballot id', 'pin', 'votes cast'], axis=1)
            workbook['File Name'] = file  # inserts file name

            # append rows to output data frame
            mergedxlsx = mergedxlsx.append(workbook, ignore_index=False)

print('Excel merge completed.')  # notifies you when merge is finished
mergedxlsx.to_excel(outputfile, index=False)  # saves output data frame to output excel file


