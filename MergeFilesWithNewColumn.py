"""Merge all xlsx files into one master spreadsheet
Insert origin source (file name) as first column"""

import os
import pandas as pd
from openpyxl import load_workbook

# copy&paste directory address of source folder NOTE: "/" at the end is important
path = r'C:\Users\jbae1\Downloads\Dummy Files\/'
# copy&paste directory address of where you want to save output file
outputfile = r'C:/Users/jbae1/OneDrive/Desktop/Output.xlsx'

# an empty data frame for new output excel file
mergedxlsx = pd.DataFrame()

# limits down to files with excel extension
for root, directories, file in os.walk(path):
    for file in file:
        if file.endswith('.xlsx'):

            df = pd.read_excel(path+file)  # reads excel file
            workbook = pd.DataFrame(df)  # assigns data frame to read excel file
            workbook['File Name'] = file  # appends new column with file name
            #print (workbook)

            # append rows to output data frame
            mergedxlsx = mergedxlsx.append(workbook, ignore_index=True)

    print('Excel merge completed.')  # notifies you when merge is finished
    mergedxlsx.to_excel(outputfile, index=False)  # saves output data frame to output excel file
