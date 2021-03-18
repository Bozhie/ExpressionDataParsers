import pandas as pd
import openpyxl
import sys

'''A module for parsing DataTracking excel files for tracking pBMF runs into ordered data for EHP set up.'''
def parseEHP(filename):

    # Pulling and reading DataTracking file
    data_tracking_file = filename
    xls = pd.ExcelFile(data_tracking_file)
    df = xls.parse('Protein and Set Number')

    # Generating top rows of standards for EHP Expression Output
    EHP_df = pd.DataFrame({'Set Number': [1, 1, 2, 2, 3, 3],
                           'Protein Number': ['noDNA', 'noDNA', 'Luciferase', 'Luciferase', 'Ubiquitin', 'Ubiquitin'],
                           'Expression Name': ['noDNA', 'noDNA', '3648_Luciferase', '3648_Luciferase', '2589_Ubiquitin', '2589_Ubiquitin']
                           })

    # Formatting data from DataTracking file to EHP output
    for i, rows in df.tail(len(df) - 3).iterrows():

        # Determine whether FSEC choice construct number is N- or C- terminally tagged variant
        if rows['FSEC choice'] == rows['N-term']:
            expression = '{}_{}_N-term'.format(rows['N-term'], rows['C part'])
        elif rows['FSEC choice'] == rows['C-term']:
            expression = '{}_{}_C-term'.format(rows['C-term'], rows['C part'])

        newdf = pd.DataFrame([[rows['Set Number'], rows['Protein Number'], expression],
                              [rows['Set Number'], rows['Protein Number'], expression]],
                             columns=list(EHP_df.columns))
        EHP_df = EHP_df.append(newdf, ignore_index=True)

    # Formatting to remove zeros
    EHP_df = EHP_df.replace(0, '-')

    # Generating output file
    output = data_tracking_file.split('.')[0] + '_EHP_output.xlsx'
    EHP_df.to_excel(output, index=False)



if __name__ == '__main__':

    # Passing in the filename
    parseEHP(sys.argv[1])