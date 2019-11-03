__author__ = 'sutha75'

import pandas as pd
import sys, os
from datetime import datetime
import re
import numpy as np
import Reconcile2 as rec

## Configs & globals ##
vbaLaunchExcelFile = '_619010_Reconciliation.xlsm'          # Excel main file to call the this RecMain script
fileNamePrefix = datetime.now().strftime("%Y-%m-%d")        # Output filename prefix
cwd = '.\\'                                                 # Default value of current working directory

# Main function
# Return exit codes:
#  0 = Normal return
#  1 = Permission error (incorrect user)
#  2 = Output file creation error. File may be open.
#  3 = VBA Excel file not found
#  4 = Input data files (SAP data, CMS journal and CMS Unallocated report) not found
#  5 = An error returned by the Reconcile() method
def main(argv):
    ##### Check tool permission #####
    currentUser = os.getlogin().lower()     # Get current Windows username and convert it to all lower case
    if currentUser == 'suyogar' or currentUser == 'sutha75':  # Allow these Windows users unrestricted access to this tool
        print('Excel reconciliation macro run by', currentUser)
    else:
        sys.exit(1)  # 1 = Permission error, set by GY.

    ##### Check program calling source and read input parameters #####
    if len(argv) == 5:  # Check if called from VBA macro
        # When called from a Excel macro, we get cwd as an argument from calling VBA macro
        cwd = argv[1] + '\\'
        try:    # Read configuration Excel file
            configData = pd.read_excel(cwd + vbaLaunchExcelFile)
            sapReportFileName = cwd + configData.iat[2, 1]
            cmsReportFileName = cwd + configData.iat[5, 1]
            cmsUnallocatedReportFileName = cwd + configData.iat[8, 1]
        except:
            print('\tFile %s does not exist. Check file name.', vbaLaunchExcelFile)
            sys.exit(3)  # Error 3 used for Config/VBA macro file read error
    else:
        # Run from command window
        cwd = '.\\'  # Direct script run using python interpreter via command window. Set current folder as cwd
        if len(argv) != 4:      # Check # of command parameters
            # Incorrect call - missing or incorrect arguments
            thisFileName = os.path.basename(__file__)
            print("\nCommand error ->\tUsage: python %s <SAP_Data.xlsx> <CMS_Journal_Report.xls> <CMS_Unallocated_Report.xls>\n" % thisFileName)
            sys.exit()
        else:
            # Command ok. Get parameters
            sapReportFileName = argv[1]
            cmsReportFileName = argv[2]
            cmsUnallocatedReportFileName = argv[3]

    print('\n\t--> Reconciling SAP data, please wait...')

    ##### Pre-process CMS Unallocated Excel data and create a formatted (duplicate receipt numbers removed and sorted by date) Excel output file #####
    try:
        df = pd.read_excel(cmsUnallocatedReportFileName)
    except:
        print(f'\tCMS Unallocated Report file name incorrect. Check {cmsUnallocatedReportFileName}')
        sys.exit(4)
    df, UnallocatedReceiptNoList = PreProcess_CmsUnallocatedReport(df)

    ##### Pre-process the SAP report Excel data and make a CSV to pass to Reconcile script #####
    try:
        df2 = pd.read_excel(sapReportFileName)
    except:
        print(f'\tSAP Report file name incorrect. Check {sapReportFileName}')
        sys.exit(4)
    df2, SapReportCsvFileName = PreProcess_SapReport(df2, UnallocatedReceiptNoList)

    ##### After generating the SAP CSV for the Reconcile script, make the following modification to df2 for final reporting using df2 data #####
    df2 = PostProcess_SapData(df2)

    ##### Pre-process the CMS Journal report and make a CSV #####
    try:
        df3 = pd.read_excel(cmsReportFileName, skiprows=2)      # Skipping top 2 unwanted rows
    except:
        print(f'\tCMS Journal Report file name incorrect. Check {cmsReportFileName}')
        sys.exit(4)
    df3, CmsReportCsvFileName = PreProcess_CmsJournal(df3)

    ##### Run the reconciliation script and get reconciled data as a dictionary #####
    try:
        sapReconciledDict = rec.Reconcile(cwd + SapReportCsvFileName, cwd + CmsReportCsvFileName)
    except Exception as ex:     # Generic exception reporting (to find out where the error occurred). Ref: https://stackoverflow.com/a/9824050/7251433
        print('Error occurred during reconciliation:')
        template = '\tAn exception of type {0} occurred. Arguments:\n{1!r}'
        message = template.format(type(ex).__name__, ex.args)
        print(message)
        sys.exit(5)     # Error code for VBA to indicate an error occurred in the Reconcile() script

    ##### Fill in the Clearing Comment column in df2 using data received from the reconciled SAP dictionary #####
    df2 = PostProcess_ReconciledData(df2, sapReconciledDict)

    ##### Generate final reconciled SAP XLSX report
    Create_Sap_Reconciled_Report(df2)

    ##### Clean up all intermediate csv and debug .txt files #####
    purgeFiles(cwd, fileNamePrefix + '_', '.txt')       # These txt files created by the Reconcile script
    purgeFiles(cwd, fileNamePrefix + '_', '.CSV')       # Two CSV files are created in this py file

    return  # Normal return from main, returns status 0 to caller
    ##### End of main #####


# Function to pre-process input CMS Unallocated Excel report data
# This function removes duplicate receipt numbers, removes time values from datatime and sorts the transaction by date.
# Returns the modified dataframe and a list of receipt numbers
def PreProcess_CmsUnallocatedReport(df):
    # Remove duplicates by looking through RECEIPT_NUMBER columns. inplace=True means df dataframe is updated with result
    df.drop_duplicates(subset=['RECEIPT_NUMBER'], keep='first', inplace=True)
    # Convert the VALUE_DATE column data, which is a string type (eg. "16/05/2018 00:00:00") to a date time object, then take only the date using dt.date method. Ref: https://stackoverflow.com/questions/32204631/how-to-convert-string-to-datetime-format-in-pandas-python
    df['VALUE_DATE'] = pd.to_datetime(df['VALUE_DATE'], format='%d/%m/%Y %H:%M:%S').dt.date
    # Sort the date frame by date ascending order
    df.sort_values(by=['VALUE_DATE'], inplace=True, ascending=True)
    Create_Formatted_Cms_Unallocated(df)
    # Get receipt numbers from column A (index 0) into a separate list. Note, the list of receipt numbers are of type integer, not string
    ReceiptNoList = df['RECEIPT_NUMBER'].tolist()
    return df, ReceiptNoList


# Function to pre-process input SAP Excel report data
# This function prepares format of certain columns and generates a CSV output file for Reconcile script to use it.
# In addition to the SAP dataframe, this function also takes UnallocatedReceiptNoList to determine if each of the SAP transation is to be marked as Unallocated or not
# Returns the modified dataframe and name of the generated CSV file
def PreProcess_SapReport(df2, UnallocatedReceiptNoList):
    # Delete unwanted data in column A
    colA = list(df2)[0]
    del df2[colA]
    # Filter off transactions that are not of type V1. How to drop pandas rows on condition? Ref: https://thispointer.com/python-pandas-how-to-drop-rows-in-dataframe-by-conditions-on-column-values/
    filteredIdx = df2[df2['Document type'] != 'V1'].index    # Filter off any transactions other than 'V1' type.
    df2.drop(filteredIdx, inplace=True)                      # Delete the rows that are not V1 type
    df2.reset_index(drop=True, inplace=True)        # Reset the index of df2, which should be missing in between indices due to data filtering above
    # Convert Document Date from datetime format to d/mm/yyyy in string format
    df2['Document Date'] = df2['Document Date'].dt.strftime('%d/%m/%Y')
    # Convert Document Number from float to int to string
    df2['Document Number'] = df2['Document Number'].astype(np.int64)
    df2['Document Number'] = df2['Document Number'].astype(str)
    # Convert Posting Key from float to int to string
    df2['Posting Key'] = df2['Posting Key'].astype(np.int64)
    df2['Posting Key'] = df2['Posting Key'].astype(str)
    # Insert a column with title 'Clearing Comment' and empty values
    val = ['' for x in range(0, len(df2))]
    df2.insert(loc=12, column='Clearing Comment', value=val, allow_duplicates=False)      # Dataframe insert method: Ref: https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.insert.html
    # Iterate through the Text field, find description text containing 'Receipt', get the receipt number and lookup if transaction is in Unallocated list
    for index, row in df2.iterrows():
        # Extract receipt number if present
        try:
            if 'Receipt' in row['Text']:    # If the word 'Receipt' in Text, extract the number (any number of digits) if found
                # Extract receipt number from Text description data and covert it to an integer, because lookup list ReceiptNoList has integers, not strings
                # Note, Regex find returns a list of strings. Since the there's only 1 receipt number, its len is always 1. So we use index 0 to get it.
                receiptNo = int( re.findall('\d+', row['Text'])[0] )
            else:
                # If there's no 'Receipt' word in the Text, then look for the usual 8-digit receipt number and get it if that exist
                receiptNo = int( re.findall('(?<!\d)\d{8}(?!\d)', row['Text'])[0] )

            # Check if the read receipt number is in the list of receipt numbers taken from CMS Unallocated data
            if receiptNo in UnallocatedReceiptNoList:
                df2.at[index, 'Clearing Comment'] = '_UNALLOCATED'    # If found in unallocated list, mark the item as Unallocated in the Clearing Comment column
        except:
            pass        # If Text field is an 'nan', the above "if 'Receipt' in.." statement will throw an error. So we catch that error here and skip it

    # Create a CSV with SAP data for Reconciliation script
    SapReportCsvFileName = fileNamePrefix + '_SAP_Report.CSV'
    df2.to_csv(cwd + SapReportCsvFileName, index=False)
    return df2, SapReportCsvFileName


# Function to post-process SAP dataframe data in preparation for the final report creation
# Returns the modified dataframe
def PostProcess_SapData(df2):
    # Add a Key column with unique numbers, as used by the Reconcile script as well
    val = [x for x in range(1, len(df2)+1)]
    df2.insert(loc=0, column='Key', value=val, allow_duplicates=False)      # Index 0 means data is inserted as the first column
    # Revert the 'Document Date' column data from string back to datetime type. Then take only the date using dt.date method. Ref: https://stackoverflow.com/questions/32204631/how-to-convert-string-to-datetime-format-in-pandas-python
    df2['Document Date'] = pd.to_datetime(df2['Document Date'], format='%d/%m/%Y').dt.date
    # Insert Receipt No column and fill it with data
    val = ['' for x in range(0, len(df2))]
    df2.insert(loc=13, column='Receipt No', value=val, allow_duplicates=False)    # Index 13 means second last column, to insert it before Clearing Comment column
    # Insert Help Comment column to help with manual clearing of unreconciled transactions
    df2.insert(loc=15, column='Sutha Help Comment', value=val, allow_duplicates=False)    # Index 15 means, insert at the end, after the Clearing Comment column
    # Iterate through the Text field, find description text containing 'Receipt', get the receipt number and fill-in data to the Receipt No column
    for index, row in df2.iterrows():
        try:
            if 'Receipt' in row['Text']:    # If the word 'Receipt' in Text, extract the number (any number of digits) if found
                # Extract receipt number from Text description data and covert it to an integer.
                df2.at[index, 'Receipt No'] = re.findall('\d+', row['Text'])[0]     # This get any numbers from Text string, and we are only using this method when Text contains the word 'Receipt'
            else:
                # If there's no 'Receipt' word in the Text, then look for the usual 8-digit receipt number and get that if it exist
                df2.at[index, 'Receipt No'] = re.findall('(?<!\d)\d{8}(?!\d)', row['Text'])[0]
        except:
            pass        # If Text field is an 'nan', the above "if 'Receipt' in.." statement will throw an error. So catch that error here and skip it

    # Iterate thru again and fill in missing Receipt number, this time if a usual receipt number found in Document Header Text column
    for index, row in df2.iterrows():  # Iterate through the Text field, find description text containing 'Receipt' and get the receipt number
        if df2.at[index, 'Receipt No'] == '':       # Fill in only if Receipt No not identified in the previous receipt number filling step
            # Extract the usual 8-digit Receipt number from 'Document Header Text'
            try:
                df2.at[index, 'Receipt No'] = re.findall('(?<!\d)\d{8}(?!\d)', row['Document Header Text'])[0]
            except:
                pass    # If Text field is an 'nan', the above "if 'Receipt' in.." statement will throw an error. So catch that error here and skip it
    return df2


# Function to pre-process input CMS Journal Excel report
# This function format the certain column data and generates a CSV for the Reconcile script
# Returns the modified dataframe and the file name of the generated CSV
def PreProcess_CmsJournal(df3):
    # Remove the empty 4th column using column index. Ref: https://stackoverflow.com/a/18145399/7251433
    df3.drop(df3.columns[4], axis=1, inplace=True)                  # axis=1 means we are dealing with a column data
    # Convert CMS date from int to string
    df3['Transaction Date'] = df3['Transaction Date'].astype(str)
    # Make CMS date length uniform by padding the missing 0 from single digit days. Ref: https://stackoverflow.com/a/31553791/7251433
    df3['Transaction Date'] = df3['Transaction Date'].str.zfill(8)
    # Create a CSV with CMS data for Reconciliation script
    CmsReportCsvFileName = fileNamePrefix + '_CMS_Report.CSV'
    df3.to_csv(cwd + CmsReportCsvFileName, index=False)
    return df3, CmsReportCsvFileName


# Function to post-process the reconciled data before in preparation for the final reconciliation report creation
# This function fixes the leading '-' issue in some description text and fills in the Clearing Comment with the information received from the Reconciled data
# Returns the modified dataframe
def PostProcess_ReconciledData(df2, sapReconciledDict):
    # Fill in the Clearing Comment column in df2 using data received from the reconciled SAP dictionary
    for index, row in df2.iterrows():
        # Also, we fix up Text description like: "- Bad Debt Write off..." to get them correctly displayed in Excel
        if row['Text'] is not np.nan:
            if row['Text'][0] == '-':                       # If first character is a '-', prefix a character space to it, so that when later opening the csv in Excel,
                df2.at[index, 'Text'] = ' ' + row['Text']   #   the description text does not turn into a formula and end up showing like #NAME? in Excel.

        # Fill in the Clearing Comment by matching the Key value. Iterate thru each item in the Reconciled dictionary
        for key, data in sapReconciledDict.items():
            if row['Key']  == key:
                if data[13] != 'No':        # Check if SAP transaction is reconciled by checking the 'Cleared' column in the Reconciled dictionary. Note: this check is now redundant as we are placing helper comments to all unreconciled trans. So there won't be any SAP trans with 'No' clearing comment
                    if data[13][0:2] == '__':   # Check if the first two characters of the comment are underscores to check if it is a helper comment
                        df2.at[index, 'Sutha Help Comment'] = data[13][2:]  # Helper comment on an unreconciled transaction. Remove the 2 leading underscores and add the comment
                    else:
                        df2.at[index, 'Clearing Comment'] = data[13]    # Normal clearing comment
    return df2


# Function to delete files with names having specific start and end pattern.
# We use this function to cleanup intermediate files such as '2019-09-16_SAP_Data_Matched.txt' and '2019-09-16_SAP_Report.CSV'
#  For the above example, we pass '2019-09-16_' as start pattern and '.txt' or '.CSV' as end pattern to clean those files.
# Ref: https://stackoverflow.com/a/5918298/7251433
def purgeFiles(dir, startPattern, endPattern):
    files = os.listdir(dir)
    for file in files:
        if file.startswith(startPattern) and file.endswith(endPattern):
            os.remove(os.path.join(dir, file))


# Function to create output xlsx file with formatted CMS Unallocated data
def Create_Formatted_Cms_Unallocated(dataFrame):
    fileName = cwd + fileNamePrefix + '_CMS_Unallocated.xlsx'     # Output file name
    try:
        with pd.ExcelWriter(fileName, engine='xlsxwriter', date_format='d/mm/yyyy') as writer:      # d/mm/yyyy means single digit day is possible as opposed to dd/mm/yyyy
            workbook = writer.book
            sheetName = 'CMS_Allocated'
            dataFrame.to_excel(writer, sheetName, index=False)
            worksheet = writer.sheets[sheetName]
            colNumFormat = workbook.add_format({'num_format': '###,###,##0.00'})
            worksheet.set_column('A:A', 17)
            worksheet.set_column('B:B', 17, colNumFormat)
            worksheet.set_column('C:C', 13)
            worksheet.set_column('D:D', 18)
            worksheet.set_column('E:E', 75)
            worksheet.set_column('F:F', 7)
            worksheet.set_column('G:G', 18)
            worksheet.set_column('H:H', 23)
            worksheet.set_column('I:I', 7.57)
            worksheet.set_column('J:J', 15)
            worksheet.set_column('K:K', 14)

            # Conditional formatting colour setup. Ref: https://xlsxwriter.readthedocs.io/example_conditional_format.html#ex-cond-format
            # Grey colour background for header titles
            formatHdr1 = workbook.add_format({'bg_color': '#A6A6A6'})

            # For header - bgnd colour applied using conditional formatting with a condition that is always expected to be True fo those cells
            worksheet.conditional_format('A1:K1', {'type': 'cell',
                                                     'criteria': '<>',
                                                     'value': 0,
                                                     'format': formatHdr1})

            # Activate autofilter on the header. Ref: https://xlsxwriter.readthedocs.io/example_autofilter.html
            worksheet.autofilter('A1:K1')
            # Freeze pane on top row
            worksheet.freeze_panes(1, 0)
    except:
        print(f'\tWrite to file \'{fileName}\' denied. Close the file and re-run script')
        sys.exit(2)

    return


# Function to create the final SAP reconciled output xlsx file with required formatting
def Create_Sap_Reconciled_Report(dataFrame):
    fileName = cwd + fileNamePrefix + '_SAP_Reconciled.xlsx'     # Output file name
    try:
        with pd.ExcelWriter(fileName, engine='xlsxwriter', date_format='d/mm/yyyy') as writer:      # d/mm/yyyy means single digit day is possible as opposed to dd/mm/yyyy
            workbook = writer.book
            sheetName = 'SAP_Reconciled'
            dataFrame.to_excel(writer, sheetName, index=False, startrow=1, header=False)
            # Startrow=1 means row 2 in Excel, header=False to use own header formatting, instead of Pandas default header format. Ref: https://xlsxwriter.readthedocs.io/example_pandas_header_format.html
            worksheet = writer.sheets[sheetName]
            colNumFormat = workbook.add_format({'num_format': '###,###,##0.00'})
            leftAlignFormat = workbook.add_format({'align': 'left'})
            rightAlignFormat = workbook.add_format({'align': 'right'})
            worksheet.set_column('A:A', 4.29, leftAlignFormat)  # Key number aligned to left
            worksheet.set_column('B:B', 11.14)
            worksheet.set_column('C:C', 12.29)
            worksheet.set_column('D:D', 3.14)
            worksheet.set_column('E:E', 4)
            worksheet.set_column('F:F', 12)
            worksheet.set_column('G:G', 3.71, rightAlignFormat) # PostKey string data is aligned to right
            worksheet.set_column('H:H', 13.57, colNumFormat)
            worksheet.set_column('I:I', 5, rightAlignFormat)    # Right align the currency column. Strings normally align to left, but we prefer on right for this one.
            worksheet.set_column('J:J', 3.43)
            worksheet.set_column('K:K', 23.71)
            worksheet.set_column('L:L', 18.71)
            worksheet.set_column('M:M', 49)
            worksheet.set_column('N:N', 10.86)
            worksheet.set_column('O:O', 46)
            worksheet.set_column('P:P', 26)

            # Add a header format
            header_format_left_aligned = workbook.add_format({
                'bold': True,
                'align': 'left',
                'valign': 'top',
                'bg_color': '#A6A6A6'})

            header_format_left_aligned_wrapped = workbook.add_format({  # Specific header format for certain column
                'bold': True,
                'align': 'left',
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#A6A6A6'})

            header_format_right_aligned_wrapped = workbook.add_format({  # Specific header format for certain column
                'bold': True,
                'text_wrap': True,
                'align': 'right',
                'valign': 'top',
                'bg_color': '#A6A6A6'})

            # Write the column headers with the defined format.  Using our own header formatting. Ref: https://xlsxwriter.readthedocs.io/example_pandas_header_format.html
            for col_num, value in enumerate(dataFrame.columns.values):
                if col_num == 2 or col_num == 5:    # Specific formatting for 'Document Number' and 'Document Date'
                    worksheet.write(0, col_num, value, header_format_left_aligned_wrapped)
                elif  col_num == 7:    # Specific formatting for 'Amount in local currency'
                    worksheet.write(0, col_num, value, header_format_right_aligned_wrapped)
                else:   # Format for all other headers
                    worksheet.write(0, col_num, value, header_format_left_aligned)

            # Cell shading formats for Clearing Comment column, so that cleared transactions are shaded as required
            reconciled_green = workbook.add_format({'bg_color': '#C6E0B4'})   # Light green
            reconciled_yellow = workbook.add_format({'bg_color': '#FFE699'})  # Light yellow
            unallocated_blue = workbook.add_format({'bg_color': '#B4C6E7'})        # Light blue
            notFoundInCms_red = workbook.add_format({'bg_color': '#DA9694'})  # Light red

            # Re-write Clearing Comment column with required cell colour shading. If conditional formatting is used, it is not easy to change colour in Excel
            for rowIdx, row in dataFrame.iterrows():  # pandas dataframe iterator
                for colIdx, cellValue in enumerate(row):
                    if colIdx != 14 and colIdx != 15:    # Skip all columns except Clearing Comment and Sutha Help Comment columns (col index 14 and 15)
                        continue
                    elif colIdx == 14:      # Clearing Comment column
                        if not cellValue:       # Checking for empty string
                            cellFormat = None
                        elif '_UNALLOCATED' in cellValue:
                            cellFormat = unallocated_blue   # Unallocated is shaded in blue
                        elif 'Verified(10)' in cellValue:
                            cellFormat = reconciled_yellow  # Verified(10) is shaded in yellow
                        else:
                            cellFormat = reconciled_green   # All others are shaded green
                        try:
                            worksheet.write_string(rowIdx+1, colIdx, cellValue, cellFormat)
                        except:
                            worksheet.write_string(rowIdx+1, colIdx, '')  # Catch any NaN, NaT cells and place a null string in those cells
                    else:   # Sutha Help Comment column
                        if 'Not found in CMS Journal' in cellValue:
                            cellFormat = notFoundInCms_red  # Shaded in light red
                        else:
                            cellFormat = None
                        try:
                            worksheet.write_string(rowIdx+1, colIdx, cellValue, cellFormat)
                        except:
                            pass  # Catch any NaN, NaT cells. Don't write anything

            # Activate autofilter on the header. Ref: https://xlsxwriter.readthedocs.io/example_autofilter.html
            worksheet.autofilter('A1:P1')
            # Freeze pane on top row
            worksheet.freeze_panes(1, 0)

    except:
        print(f'\tWrite to file \'{fileName}\' denied. Close the file and re-run script')
        sys.exit(2)

    return


if __name__ == "__main__":
    main(sys.argv)
