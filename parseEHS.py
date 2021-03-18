import pandas as pd
import openpyxl
import sys

'''A module for parsing DataTracking excel files for tracking pBMF runs into ordered data for EHS input.'''
def parseEHS(filename):

    # Pulling and reading DataTracking file
    data_tracking_file = filename
    xls = pd.ExcelFile(data_tracking_file)
    df = xls.parse('Protein and Set Number')

    # Generating beginning standards of EHS Expression Output
    EHS_df = pd.DataFrame({'Set Number': [1, 2, 3], 'Protein Number': ['noDNA', 'Luciferase', 'Ubiquitin'],
                           'Expression Name': ['noDNA', '3648_Luciferase', '2589_Ubiquitin'],
                           'Construct': [0, 3648, 2589]})

    # Formatting data from DataTracking file to EHS output
    for i, rows in df.tail(len(df) - 3).iterrows():
        N_expression = '{}_{}_N-term'.format(rows['N-term'], rows['C part'])
        C_expression = '{}_{}_C-term'.format(rows['C-term'], rows['C part'])
        newdf = pd.DataFrame([[rows['Set Number'], rows['Protein Number'], N_expression, rows['N-term']],
                              [rows['Set Number'], rows['Protein Number'], C_expression, rows['C-term']]],
                             columns=list(EHS_df.columns))
        EHS_df = EHS_df.append(newdf, ignore_index=True)

    # Adding tail set of standards for EHS
    footer = pd.DataFrame({'Set Number': [0, 0, 0], 'Protein Number': ['noDNA', 'Luciferase', 'Ubiquitin'],
                           'Expression Name': ['NoDNA', '3648_Luciferase', '2589_Ubiquitin'], 'Construct': [0, 0, 0]})
    EHS_df = EHS_df.append(footer, ignore_index=True)

    # Formatting to remove zeros
    EHS_df = EHS_df.replace(0, '-')

    # Generating output file
    output = data_tracking_file.split('.')[0] + '_EHS_output.xlsx'
    EHS_df.to_excel(output, index=False)


if __name__ == '__main__':

    parseEHS(sys.argv[1])