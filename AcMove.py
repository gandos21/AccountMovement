__author__ = 'sutha75'

import pandas as pd
import sys, os
from datetime import datetime
import threading
import PySimpleGUI as sg
import json
import re

## Configs & globals ##
ReportFileName = datetime.now().strftime("%Y-%m-%d") + '_Account_Movement.xlsx'     # Output file name
CallCounter = 0

# Main function
def main(argv):
    ## Read configuration Excel file
    if isinstance(argv, list):
        guiCall = False             # Called from command line
        try:
            movementDataFile = argv[1]
        except:
            print('\tERROR: Account movement Excel file invalid!')
            return
    else:
        guiCall = True              # Called from GUI panel
        movementDataFile = argv['-FileName-'].Get()     # Get the value from FileName input text box GUI element

    ## Read the Movement input data file
    try:
        df = pd.read_excel(movementDataFile, skiprows=1)
    except:
        print('--> ERROR: Account movement Excel file name incorrect!')
        if guiCall:
            argv.write_event_value('-Thread Done-', '')     # Notify GUI that the thread has finished as we are about to return. This is not necessary, but we use this callback method to allow subsequent thread starts
        return

    print('--> Balancing movement data, please wait...')

    ## Add ABS formulas to Absolute Amount column
    rowIdx = 3        # Data start row (when viewed in Excel). i.e. the actual first data row, not the title row.
    col_C_Data = []
    for i in range(len(df)):
        # Required example formula for column C: '=ABS(B3)'
        col_C_Data.append(f'=ABS($B{str(rowIdx)})')
        rowIdx += 1
    df['Absolute Amount'] = col_C_Data

    ## Add a column with title 'Cleared' and load all rows with 'No' values
    df['Cleared'] = 'No'

    ## Set Cleared flag to Yes for all items with zero amount
    #   Setting a column value based on value in another column. Ref: https://stackoverflow.com/a/49161313/7251433
    #   Combining multiple conditions with &-operator.           Ref: https://stackoverflow.com/a/15315507/7251433
    df.loc[(df['movement total'] > -0.005) & (df['movement total'] < 0.005), 'Cleared'] = 'Yes'     # Because the amount is a float, we use relative comparison to check for near zero

    ## Iterate through the dataframe, find matching amount pairs and clear them
    if guiCall == False:
        print('\n\tAmount balancing in progress: ', end='', flush=True) # Printing w/o a newline and flushing it to immediately appear on screen. Ref: https://stackoverflow.com/a/493399/7251433

    # Before starting the double loop, get the total number of uncleared items for progress % calculation
    xStart = xPrev = df.loc[df['Cleared'] == 'No', 'Cleared'].count()
    y = z = 0
    for index, row in df.iterrows():
        if df.at[index, 'Cleared'] == 'No':  # Fill in only those items that are not yet cleared

            # Monitor remaining uncleared items for progress % calculation
            x = df.loc[df['Cleared'] == 'No', 'Cleared'].count()    # Get current remaining uncleared items anc check if it has reduced
            if xPrev != x:
                xPrev = x
                y += 2          # If we cleared a pair in the previous loop, then we add 2 for cleared pair
            else:
                y += 1          # Otherwise increment 1 for each outer loop pass. Even if don't clear all, the progress % has to be 100% when outer loop ends
            z += 1              # z counter simply is used to limit GUI progress value notification and nothing to do this progress value calculation itself

            # For each item, find a matching amount. Amounts are both +ve and -ve. So we add two of them to check if they sum to zero, and then clear that pair
            for index2, row2 in df.iterrows():
                if df.at[index2, 'Cleared'] == 'No':    # Again, we only consider uncleared items
                    if IsFloatValueZero( df.at[index, 'movement total'] + df.at[index2, 'movement total'] ):
                        df.at[index,  'Cleared'] = 'Yes'
                        df.at[index2, 'Cleared'] = 'Yes'
                        break

            # Update progress %, since the this double loop can take sometime to finish. We calculate progress and report from the outer loop
            progress = round((float(y) / xStart) * 100)             # When round() is used without digits paramater, it will return a rounded up integer value. Ref: https://docs.python.org/3/library/functions.html#round
            if guiCall == False:        # Progress printing for command prompt
                if progress < 100:      # Progress value may be on one value repeated times, eg. 100 may come round a few times due to round() function used above. So we do the printing for 100% separately. If we don't use round() function and int() instead, we may not see 100, due to 99.999999 problem
                    print('% 3d%%' % progress, end='', flush=True)  # Printing w/o a newline and flushing it to immediately appear on screen. Ref: https://stackoverflow.com/a/493399/7251433
                else:
                    print('%d%%' % progress, end='', flush=True)
                print('\b\b\b\b', end='', flush=True)       # Move back the cursor 4 positions, so that the progress xxx% shown on screen appears in the same location
            else:
                # Update progress value to GUI panel by issuing an event
                if z % 3 == 0 or progress == 100:            # Limit GUI progress to every 3 outer loop pass or if the progress value has reached 100. If we don't allow for 100, GUI progressbar may stop short of fully completing as we skip 2 of 3 loop
                    argv.write_event_value('-Progress Value-', progress)    # Feedback from this thread to GUI. Ref: https://pysimplegui.readthedocs.io/en/latest/cookbook/#recipe-long-operations-multi-threading

    print('\n--> Number of amount checks performed: {:,d}'.format(CallCounter))

    ## Generate final movement XLSX report
    try:
        Create_Movement_Report(df)
    except:
        print('\n--> ERROR: Write to Excel output file denied. File may be open!')
    else:
        print('\n--> Success: Balancing task completed! Check generated Excel report.')

    if guiCall:
        argv.write_event_value('-Thread Done-', '')  # Notify GUI that the thread has finished as we are about to return. This is not necessary, but we use this callback method to allow subsequent thread starts
    return
    ##### End of main() #####


# Function to check if a given value of type float is near zero
def IsFloatValueZero(floatValue):
    global CallCounter
    CallCounter += 1
    if floatValue < 0.001 and floatValue > -0.001:
        return True
    else:
        return False


# Function to create the amount movement output xlsx file with required formatting
def Create_Movement_Report(dataFrame):
    try:
        with pd.ExcelWriter(ReportFileName, engine='xlsxwriter', date_format='d/mm/yyyy') as writer:      # d/mm/yyyy means single digit day is possible as opposed to dd/mm/yyyy
            workbook = writer.book
            sheetName = 'Account movement'
            dataFrame.to_excel(writer, sheetName, index=False, startrow=2, header=False)
            # Startrow=2 means row 3 in Excel, header=False to use own header formatting, instead of Pandas default header format. Ref: https://xlsxwriter.readthedocs.io/example_pandas_header_format.html
            worksheet = writer.sheets[sheetName]
            amountFormat = workbook.add_format({'num_format': '$###,###,##0.00'})
            centreAlignFormat = workbook.add_format({'align': 'center'})
            worksheet.set_column('A:A', 18.43, centreAlignFormat)
            worksheet.set_column('B:B', 19.29, amountFormat)
            worksheet.set_column('C:C', 19.29, amountFormat)
            worksheet.set_column('D:D', 15.86)
            worksheet.set_column('F:F', 40.00)      # Cell F1 is used to print a summary note if there are any unmatched amounts found

            # Conditional format used for summary note in cell F1, so when the note is deleted in Excel, the formating will disapper
            formatF1 = workbook.add_format({'bg_color': '#F4B084'})
            formatF1.set_bold()
            worksheet.conditional_format('F1', {'type': 'cell', 'criteria': '<>', 'value': 0, 'format': formatF1})

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
            for col_num, value in enumerate(dataFrame.columns.values):
                if col_num == 0 or col_num == 1:    # Specific formatting for 'contract number' and 'movement total'
                    worksheet.write(rowNum, col_num, value, header_format_centre_aligned_green)
                elif col_num == 2:   # 'Absolute Amount'
                    worksheet.write(rowNum, col_num, value, header_format_centre_aligned_blue)
                else:   # Comment (Cleared) column
                    worksheet.write(rowNum, col_num, 'Comment', header_format_centre_aligned_blue)  # Using title 'Comment' instead of 'Cleared'
            worksheet.set_row(rowNum, 24.75)             # Set row height for the header row

            # Cell shading formats for items that are not cleared
            uncleared_ContractNo = workbook.add_format({'bg_color': '#F4B084', 'align': 'center'})  # F4B084 is Terracotta shade
            uncleared_Amount = workbook.add_format({'bg_color': '#F4B084', 'num_format': '$###,###,##0.00'})
            uncleared_Comment = workbook.add_format({'bg_color': '#F4B084'})
            cleared_Comment = workbook.add_format({'bg_color': '#C4D79B'})      # C4D79B is Light green shade for cleared comments

            # Re-write data with required cell colour shading for items that are not cleared.
            # Note, we are not re-writing everything here, but only the cells that require shading. Unshaded (Cleared) cells are filled in by pd function to_excel() above
            rowOffset = 2
            for rowIdx, row in dataFrame.iterrows():
                if dataFrame.at[rowIdx, 'Cleared'] == 'No':  # Check if item is uncleared
                    for colIdx, cellValue in enumerate(row):
                        if colIdx == 0:                     # 'contract number' column
                            worksheet.write(rowIdx+rowOffset, colIdx, cellValue, uncleared_ContractNo)  # Write() function Ref: https://xlsxwriter.readthedocs.io/worksheet.html
                        elif colIdx == 1 or colIdx == 2:    # 'movement total' and 'Absolute Amount'
                            worksheet.write(rowIdx+rowOffset, colIdx, cellValue, uncleared_Amount)
                        else:   # 'Cleared' column data
                            worksheet.write(rowIdx+rowOffset, colIdx, 'Unmatched', uncleared_Comment)   # Replacing 'No' with 'Unmatched'
                else:   # Cleared item - green shade the Comment data
                    for colIdx, cellValue in enumerate(row):
                        if colIdx == 3:     # 'Cleared' column
                            worksheet.write(rowIdx+rowOffset, colIdx, 'Ok', cleared_Comment)            # Instead of using 'Yes', we are using 'Ok' for the comment

            # Print the summary note with uncleared information in cell F1
            unclearedCount = dataFrame.loc[dataFrame['Cleared'] == 'No', 'Cleared'].count()     # Counting the number of uncleared items. How many rows with Cleared==No?  Ref: https://stackoverflow.com/a/46080857
            if unclearedCount > 0:
                worksheet.write('F1', 'Found ' + str(unclearedCount) + ' contracts with unmatched amounts!')    # No cell formatting applied here, because we have already setup a conditional formatting for cell F1 above
            # Activate autofilter on the header. Ref: https://xlsxwriter.readthedocs.io/example_autofilter.html
            worksheet.autofilter('A2:D2')
            # Freeze pane on 2nd row
            worksheet.freeze_panes(2, 0)    # Freeze pane set on cell A3
    except:
        raise Exception     # Unable to write to Excel output file - file may be open. Raise exception back to the caller
    return


# GUI main
def mainGUI():
    global CallCounter

    panelDefaultsFileName = 'PanelDefaults.json'
    panelDefaults = {'-FileName-' : os.getcwd() + '\\_AcMovement.xls'}

    # Save the sys.stdout object pointer. With the window.read() call below, the stdout pointer will switch to the GUI output panel, because of the added sg.Output() element in the GUI layout. Ref: https://pysimplegui.readthedocs.io/en/latest/#output-element
    #  So any print() calls after that will appear on GUI only. Hence, we save the original stdout object poitner to print to command window for debugging purpose.  Ref: https://stackoverflow.com/a/3263733
    cmdOut = sys.stdout
    cmdErr = sys.stderr

    # Setting window colour theme
    sg.theme('SandyBeach')
    #print(sg.theme_list())     # Debug: Print all available themes in PySimpleGUI. Also check at, https://pysimplegui.readthedocs.io/en/latest/#themes-automatic-coloring-of-your-windows

    # Load panel value defaults from json
    try:  # If file exist load from file, otherwise (file doesn't exist or doesn't contain expected data) create a new json
        with open(panelDefaultsFileName, 'r') as fp:
            panelDefaults = json.load(fp)                       # Storing/restoring dictionary to/from a json
    except:
        with open(panelDefaultsFileName, 'w') as fp:
            json.dump(panelDefaults, fp, indent=4)              # Creating with indentation. Ref: https://stackoverflow.com/a/12309296
        os.system(f'attrib +h {panelDefaultsFileName}')         # Set as a hidden file. Ref: https://stackoverflow.com/a/58016586

    # Define panel layout using default values read from json or initial defaults
    panel_layout = [
        # Title text
        [sg.Text('Account Movement Balancer', font='Any 15', pad=(5,(8,15)))],

        # File name input and browse trigger
        [sg.Text('Movement Statement', size=(15,1), pad=(5,(3,15))),
         sg.InputText(panelDefaults['-FileName-'], size=(60,1), key='-FileName-', enable_events=True, pad=(5,(3,15))),
         sg.FileBrowse(initial_folder=os.getcwd(), key='-FileBrowse-', file_types=(('Excel Files', ('*.xlsx', '*.xls')),('All Files', '*.*')), pad=(5,(3,15)))   # Browsing with init folder Ref: https://github.com/PySimpleGUI/PySimpleGUI/issues/239
        ],

        # Output element inside a frame
        [ sg.Frame('Messages',[ [sg.Output(size=(83,8), key='-Output-')] ]) ],

        # Progress bar element. Ref: https://pysimplegui.readthedocs.io/en/latest/#progressbar-element
        [sg.Text('Progress')],
        [sg.ProgressBar(100, orientation='h', size=(47,20), key='-Progressbar-', pad=(5,(3,15)))],

        # Buttons
        [sg.Button('Start',              size=(14,2), font='Any 12', pad=((8,5), 3)),  # Pixel padding is used to fine tune around any element.
         sg.Button('Open Excel Report',  size=(17,2), font='Any 12'),                  # ((left,right), (top,bottom)) Default: 5pixels on x-axis and 3pixels on y-axis. Ref: https://pysimplegui.readthedocs.io/en/latest/#pad
         sg.Button('Remove Old Reports', size=(17,2), font='Any 12'),
         sg.Button('Exit',               size=(14,2), font='Any 12')
        ]
    ]

    # Create the window object
    window = sg.Window('Account Movement', panel_layout, default_element_size=(80, 1), grab_anywhere=False)
    # Using a thread to run main() as a separate process, so this GUI doesn't freeze up due to long runtime required in main(), which takes more than 15s to complete its operation
    #  Multi-threading reference: https://pysimplegui.readthedocs.io/en/latest/cookbook/#recipe-long-operations-multi-threading
    #    Threading doc: https://docs.python.org/3/library/threading.html
    #    Daemon vs non-daemon threads: https://stackoverflow.com/a/190017
    mainThread = None
    valuesCopy = panelDefaults                  # Initial state of window elements' values will be panelDefaults

    # Main event handler loop
    while True:
        event, values = window.read()           # Read event from window. Buttons are event enabled. Events for other elements enabled (using parameter enable_events) as desired. Ref: https://pysimplegui.readthedocs.io/en/latest/#events
        ## Exit button or window close (X) event ##
        if event in (sg.WIN_CLOSED, 'Exit'):    # Checking for window X close button or our own Exit button. Checking of X is prioritised over other events. Doing X abruptly stop compiled EXE execution, eg. doing json dump above this line was crashing compiled EXE when X was clicked.
            break                               #  When X is clicked to close, window object will return None values in 'values', i.e. no dictionary values. event will be None too.
        #print(event, ' --> ', values)          # Debug print
        ## Button events ##
        if event == 'Start':
            if mainThread == None:              # Check we can create and start a thread. We don't want to create one when is one active. Note, a thread can only be started once, else gives runtime error. Once the thread run is finished, new one has to be created.
                CallCounter = 0                 # Global CallCounter is used to count # of amount checks. Clear for each run.
                window.FindElement('-Output-').Update('')   # Clearing the contents of Output element window. Ref: https://github.com/PySimpleGUI/PySimpleGUI/issues/1441#issuecomment-493741474
                mainThread = threading.Thread(target=main, args=(window,), daemon=True)     # main() set as target, window set as main's parameter
                mainThread.start()              # Start main() thread
            elif mainThread.is_alive():         # In case the Start button is clicked multiple while we have an instance of main() thread running, reject subsequent request until thread run is done
                print('Script already running!')
            else:
                pass
        if event == 'Open Excel Report':
            if os.path.isfile(ReportFileName):
                sysCommand = f'start \"excel\" \"{os.getcwd()}\\{ReportFileName}\"'
                os.system(sysCommand)               # Launching Excel to open the config file. Note: Excel requires double quote for its parameters. Ref: https://stackoverflow.com/a/57948434
            else:
                print(f'ERROR: Excel report {ReportFileName} does not exist! Start script to create it.')
        if event == 'Remove Old Reports':
            # Regex '(20\d\d-\d\d-\d\d_Account_Movement\.xlsx)(?<!^2018-08-02_Account_Movement\.xlsx$)' was built at https://regex101.com/r/FV9B3a/1
            #  It is to capture all Excel report files except for the current generated file. '?<!' is a negative lookback that excludes the given string, ^ and $ marks start and end of string.
            #  In the pattern created, I have inserted marker '___FILENAME___' to replace it with acutal filename, so we generate a dynamic regex pattern based on current report file name.
            regPattern = '(20\\d\\d-\\d\\d-\\d\\d_Account_Movement\\.xlsx)(?<!^___FILENAME___$)'.replace('___FILENAME___',  ReportFileName.replace('.', '\\.'))     # Also replacing . with \. so file extension dot is taken as literal in regex
            try:
                FilePurge('.\\', regPattern)
            except:
                print('ERROR: Could not remove old report files! File may be open.')
            else:
                print('Success: Old report files removed')

        ## Thread call-back events ##
        if event == '-Thread Done-':            # Call back event from main() to indicate thread run completed
            mainThread = None                   #  Clear the thread handler - ready to accept another Start
        if event == '-Progress Value-':         # Call back event from main() to with main script progress value
            window['-Progressbar-'].Update(values['-Progress Value-'])        # Update current value to progressbar

        ## File name input text box change events ##
        if event == '-FileName-':                    # File name updated
            values['-FileName-'] = values['-FileName-'].replace('/', '\\')  # FileBrowse puts forward slash only. Replace them with back slash
            window['-FileName-'].Update(values['-FileName-'])               # Using Update() to update the value of the InputText box.  Ref: https://pysimplegui.readthedocs.io/en/latest/call%20reference/#window/#Update

        # If any window element values changed, backup the values
        if values != valuesCopy:  # We compare against a saved double copy and only update json when there's a difference.  Comparing two dictionaries  Ref: https://stackoverflow.com/a/40921229
            valuesCopy = values   # We do this double backup of values, because when clicking the X button to close the window will yield a None in 'values'. Later when writing to json, we use valuesCopy, which will have valid data

    # Closing the GUI window after Exit button or window X is clicked
    window.close()
    sys.stdout = cmdOut     # Restore stdout and stderr since Output GUI element had changed those object pointers to tkinter output
    sys.stderr = cmdErr

    # Save the window element values to json file, if values changed from last saved data
    #  valuesCopy may contain additional key/value pair due to FileBrowse and other call-back events. Take only -FileName- and drop the rest
    valuesCopy = {k:v for k,v in valuesCopy.items() if k in ['-FileName-']}
    if valuesCopy != panelDefaults:
        os.system(f'attrib -h {panelDefaultsFileName}')     # We have set the json as a hidden file when initially created. We need to unhide to edit, otherwise there will be permission denied error. Ref: https://stackoverflow.com/a/63697862/7251433
        with open(panelDefaultsFileName, 'w') as fp:
           json.dump(valuesCopy, fp, indent=4)
        os.system(f'attrib +h {panelDefaultsFileName}')     # Make it hidden again
        print('\t->JSON file updated')          # This print will go to command prompt window, since we have now closed GUI window and restored original stdout pointer


# Function to delete group of files using regex pattern. Ref: https://stackoverflow.com/a/1548720
# re.search() changed to re.match() as we are trying to match full string found in directory, not a partial match. search() vs match() Ref: https://stackoverflow.com/a/180993
def FilePurge(dir, pattern):
    for f in os.listdir(dir):
        if re.match(pattern, f):
            os.remove(os.path.join(dir, f))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        main(sys.argv)      # If there is a parameter given on command line, run main() with that.  eg. C:\>python AcMove.py MovementFile.xls
    else:
        mainGUI()           # Call the GUI if called with no parameter eg. C:\>python AcMove.py
