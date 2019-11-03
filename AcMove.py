__author__ = 'sutha75'

import pandas as pd
import sys, os
from datetime import datetime
from time import sleep

## Configs & globals ##
vbaLaunchExcelFile = '_AcMove_Macro.xlsm'               # Excel main file to call the this RecMain script
fileNamePrefix = datetime.now().strftime("%Y-%m-%d")    # Output filename prefix
callCntr = 0


# Main function
# Return exit codes:
#  0 = Normal return
#  1 = Permission error (incorrect user)
#  2 = Output file creation error. File may be open.
#  3 = VBA Excel file not found
#  4 = Input data files (Movement input data file) not found
#  5 = Not used
def main(argv):
    ## Check tool permission
    currentUser = os.getlogin().lower()     # Get current Windows username and convert it to all lower case
    if currentUser == 'suyogar' or currentUser == 'sutha75':  # Allow these Windows users unrestricted access to this tool
        print('Account Movement macro run by', currentUser)
    else:
        sys.exit(1)  # 1 = Permission error, set by GY.

    ## Check program calling source and read input parameters
    if len(argv) == 5:  # Check if called from VBA macro
        # When called from a Excel macro, we get cwd as an argument from calling VBA macro
        cwd = argv[1] + '\\'
        print(f'VBA launch path: {cwd}')
    else:
        # Script is run from command window
        cwd = '.\\'  # Direct script run using python interpreter via command window. Set current folder as cwd
        if len(argv) != 1:      # Check # of command parameters
            # Incorrect call - missing or incorrect arguments
            thisFileName = os.path.basename(__file__)
            print('\nCommand error ->\tUsage: python %s \n' % thisFileName)
            sys.exit()

    ## Read configuration Excel file
    try:
        configData = pd.read_excel(cwd + vbaLaunchExcelFile)
        movementDataFile = cwd + configData.iat[2, 1]
    except:
        print('\tFile %s does not exist. Check file name.', vbaLaunchExcelFile)
        sys.exit(3)  # Error 3 used for Config/VBA macro file read error

    print('\n\t--> Balancing movement data, please wait...')

    ## Read the Movement input data file
    try:
        df = pd.read_excel(movementDataFile, skiprows=1)
    except:
        print(f'\tAccount Movement file name incorrect. Check {movementDataFile}')
        sys.exit(4)

    ## Add ABS formulas to Absolute Amount column
    rowIdx = 3        # Data start row (when viewed in Excel). i.e. the actual first data row, not the title row.
    col_C_Data = []
    for i in range(len(df)):
        # Required example formula for column C: '=ABS(B3)'
        col_C_Data.append(f'=ABS($B{str(rowIdx)})')
        rowIdx += 1
    df['Absolute Amount'] = col_C_Data

    ## Add a column with title 'Clearing Comment' and load with 'No' values
    val = ['No' for i in range(len(df))]
    df['Cleared'] = val

    ## Set Cleared flag to Yes for all items with zero amount
    #   Setting a column value based on value in another column. Ref: https://stackoverflow.com/a/49161313/7251433
    #   Combining multiple conditions with &-operator.           Ref: https://stackoverflow.com/a/15315507/7251433
    df.loc[(df['movement total'] > -0.001) & (df['movement total'] < 0.001), 'Cleared'] = 'Yes'     # Because the amount is a float, we use relative comparison to check for near zero

    ## Iterate through the dataframe, find matching amount pairs and clear them
    print('\n\tAmount balancing in progress: ', end='', flush=True) # Printing w/o a newline and flushing it to immediately appear on screen. Ref: https://stackoverflow.com/a/493399/7251433
    for index, row in df.iterrows():
        if df.at[index, 'Cleared'] == 'No':  # Fill in only those items that are not yet cleared
            # For each item, find a matching amount. Amounts are both +ve and -ve. So we add two of them to check if they sum to zero, and then clear that pair
            for index2, row2 in df.iterrows():
                if df.at[index2, 'Cleared'] == 'No':    # Again, we only consider uncleared items
                    if IsFloatValueZero( df.at[index, 'movement total'] + df.at[index2, 'movement total'] ):
                        df.at[index, 'Cleared'] = 'Yes'
                        df.at[index2, 'Cleared'] = 'Yes'
                        break
        # Print progress %, since the above double loop can take sometime to finish
        progress = round((float(index) / len(df)) * 100)
        if progress < 100:      # Progress value may be on one value repeated times, eg. 100 may come round a few times due to round() function used above. SO we do the printing for 100% separately. If don't use round() function and int() instead, we may not see 100, due to 99.999999 problem
            print('% 3d%%' % progress, end='', flush=True)  # Printing w/o a newline and flushing it to immediately appear on screen. Ref: https://stackoverflow.com/a/493399/7251433
        else:
            print('%d%%' % progress, end='', flush=True)
        print('\b\b\b\b', end='', flush=True)       # Move back the cursor 4 positions, so that the progress xxx% shown on screen appears in the same location
    print('\nNumber of amount comparisons performed during the balancing process = {:,d}'.format(callCntr))

    ## Debug outputs
    #print(df)
    #print(df['Abs Amount'].sum())  # Print sum of the amount column, just to cross check with actual data in Excel
    #df.to_csv('Z_Test_Out.csv', sep=',')

    ## Generate final movement XLSX report
    Create_Movement_Report(cwd, df)
    sleep(1)    # Wait for 1s before ending

    return  # Normal return from main, returns status 0 to caller
    ##### End of main #####


# Function to check if a given value of type float is near zero
def IsFloatValueZero(floatValue):
    global callCntr
    callCntr += 1
    if floatValue < 0.001 and floatValue > -0.001:
        return True
    else:
        return False


# Function to create the amount movement output xlsx file with required formatting
def Create_Movement_Report(cwd, dataFrame):
    # Remove the Cleared column from dataframe, before writing it to XLSX. We don't want this column to appear in the output
    df2 = dataFrame.drop('Cleared', 1)      # Deleting a data column. Ref: https://stackoverflow.com/a/18145399/7251433
    fileName = cwd + fileNamePrefix + '_Account_Movement.xlsx'     # Output file name
    try:
        with pd.ExcelWriter(fileName, engine='xlsxwriter', date_format='d/mm/yyyy') as writer:      # d/mm/yyyy means single digit day is possible as opposed to dd/mm/yyyy
            workbook = writer.book
            sheetName = 'Account movement'
            df2.to_excel(writer, sheetName, index=False, startrow=2, header=False)
            # Startrow=2 means row 3 in Excel, header=False to use own header formatting, instead of Pandas default header format. Ref: https://xlsxwriter.readthedocs.io/example_pandas_header_format.html
            worksheet = writer.sheets[sheetName]
            amountFormat = workbook.add_format({'num_format': '$###,###,##0.00'})
            centreAlignFormat = workbook.add_format({'align': 'center'})
            worksheet.set_column('A:A', 18.43, centreAlignFormat)
            worksheet.set_column('B:B', 19.29, amountFormat)
            worksheet.set_column('C:C', 19.29, amountFormat)

            # Add row header formats
            header_format_centre_aligned_green = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'top',
                'bg_color': '#9ACD32'})

            header_format_centre_aligned_blue = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'top',
                'bg_color': '#B4C6E7'})

            # Print worksheet header
            rowNum = 0
            cellFormat = workbook.add_format()
            cellFormat.set_font_size(20)
            cellFormat.set_font_color('#006400')    # Green
            cellFormat.set_bold()
            worksheet.write('A1', 'Account movement', cellFormat)
            worksheet.set_row(rowNum, 25.5)  # Set row height for the header row

            # Write the column headers with the defined format.  Using our own header formatting. Ref: https://xlsxwriter.readthedocs.io/example_pandas_header_format.html
            rowNum = 1
            for col_num, value in enumerate(df2.columns.values):
                if col_num == 0 or col_num == 1:    # Specific formatting for 'contract number' and 'movement total'
                    worksheet.write(rowNum, col_num, value, header_format_centre_aligned_green)
                else:   # 'Absolute Amount' column
                    worksheet.write(rowNum, col_num, value, header_format_centre_aligned_blue)
            worksheet.set_row(rowNum, 24.75)             # Set row height for the header row

            # Cell shading formats for items that are not cleared
            uncleared_ContractNo = workbook.add_format({'bg_color': '#F4B084', 'align': 'center'})  # F4B084 is Terracotta shade
            uncleared_Amount = workbook.add_format({'bg_color': '#F4B084', 'num_format': '$###,###,##0.00'})

            # Re-write data with required cell colour shading for items that are not cleared
            rowOffset = 2
            for rowIdx, row in dataFrame.iterrows():
                if dataFrame.at[rowIdx, 'Cleared'] == 'No':  # Check if item is uncleared
                    for colIdx, cellValue in enumerate(row):
                        if colIdx == 0:                     # 'contract number' column
                            worksheet.write(rowIdx+rowOffset, colIdx, cellValue, uncleared_ContractNo)      # Write() function Ref: https://xlsxwriter.readthedocs.io/worksheet.html
                        elif colIdx == 1 or colIdx == 2:    # 'movement total' and 'Absolute Amount'
                            worksheet.write(rowIdx+rowOffset, colIdx, cellValue, uncleared_Amount)
                        else:   # Ignore the 'Cleared' column data
                            pass

            # Activate autofilter on the header. Ref: https://xlsxwriter.readthedocs.io/example_autofilter.html
            worksheet.autofilter('A2:C2')
            # Freeze pane on 2nd row
            worksheet.freeze_panes(2, 0)

    except Exception as ex:  # Generic exception reporting (to find out where the error occurred). Ref: https://stackoverflow.com/a/9824050/7251433
        template = '\tAn exception of type {0} occurred. Arguments:\n{1!r}'
        message = template.format(type(ex).__name__, ex.args)
        print(message)
        print(f'\tWrite to file \'{fileName}\' denied. Close the file and re-run script')
        sys.exit(2)
    return


if __name__ == "__main__":
    main(sys.argv)
