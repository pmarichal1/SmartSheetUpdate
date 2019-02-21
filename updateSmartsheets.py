'''
Copyright (c) 2019

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''
from __future__ import print_function
from __future__ import unicode_literals
from __future__ import absolute_import
import sys
import csv
import platform
#import logging
import smartsheet
from smartsheet import sheets
import time
import colorama

# hack to fix problem with reading Unicode from smartsheets
if sys.version_info.major < 3:
    import sys
    reload(sys)
    sys.setdefaultencoding('utf8')
    print("Fixing Unicode issue with python 2.x")

class Bcolors(object):
    '''color pallette for print statements'''
    HEADER = '\033[96m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'


def get_access_token(program_input_path):
    '''go get the access token we need for smartsheet SDK request'''
    try:
        token_hndl = open(program_input_path + "accessToken.txt", 'r')
        access_token = token_hndl.readline()
        token_hndl.close()
        return access_token
    except IOError:
        print(Bcolors.FAIL + "No such file:" +
              program_input_path + "accessToken.txt" + Bcolors.ENDC)
        sys.exit(1)


def get_smartsheet_ids(program_input_path):
    '''# go get the Smartsheet IDs we want to download'''
    try:
        ss_hndl = open(program_input_path + "smartsheetUpdateIDs.txt", 'r')
        line_list = ss_hndl.readlines()
        # create list of IDs
        ids_list = []
        for line in line_list:
            ids_list.append(line.strip())
        ss_hndl.close()
        return ids_list
    except IOError:
        print(Bcolors.FAIL + "No such file:" +
              program_input_path + "smartsheetUpdateIDs.txt" + Bcolors.ENDC)
        sys.exit(1)


def read_jira_import_file(jira_input_path):
    '''create a tuple list from the JIRAInputData file'''
    try:
        with open(jira_input_path + 'JIRAImportData.txt', mode='r') as jira_hdnl:
        #with open(jira_input_path + 'JIRAImportData.txt', encoding='ISO-8859-1', mode='r') as jira_hdnl:
            reader = csv.reader(jira_hdnl, delimiter='\t')
            if sys.version_info.major < 3:
                jira_list = map(tuple, reader)
            else:
                jira_list = list(map(tuple, reader))
            jira_hdnl.close()
            return jira_list
    except IOError:
        print(Bcolors.FAIL + "No such file:" + jira_input_path + 'JIRAImportData.txt' + Bcolors.ENDC)
        sys.exit(1)


def get_cell_by_column_name(row_id, column_name, column_map):
    '''Helper function to find cell in a row and return name'''
    try:
        # make sure the key we are looking up exists
        column_id = column_map[column_name]
        if row_id.get_column(column_id) == None:
            print(Bcolors.FAIL + f"Column name '{column_name}' returns None" + Bcolors.ENDC)
        return row_id.get_column(column_id)
    except KeyError:
        print(Bcolors.FAIL + f"Column name '{column_name}' does not exist" + Bcolors.ENDC)
        return 0


def evaluate_row_and_build_updates(ss_client, source_row, jira_entry, column_map):
    '''Find the cell and value we want to evaluate'''
    #   for "Input into JIRA" value from the cell
    cell_data_checkbox = get_cell_by_column_name(source_row, "Input into JIRA", column_map)
    if cell_data_checkbox == None or cell_data_checkbox == 0:
        return None, 0
    checkbox = cell_data_checkbox.value
    if checkbox != True:  # Skip if already true
        # get contents of description column
        cell_data_description = get_cell_by_column_name(
            source_row, "Additional description if necessary", column_map)
        # get "Additional description if necessary" value from cell
        if cell_data_description == None or cell_data_description == 0:
            return None, 0
        description = cell_data_description.display_value

        # get "JOB #" value from cell
        cell_data_job = get_cell_by_column_name(source_row, "Job #", column_map)
        if cell_data_job == None or cell_data_job == 0:
            return None, 0
        primary = cell_data_job.display_value

        # get "chapter #s" value from cell
        cell_data_chapter = get_cell_by_column_name(source_row, "Chapter # or #s", column_map)
        if cell_data_chapter == None or cell_data_chapter == 0:
            return None, 0
        chapter = cell_data_chapter.display_value

        # get "Summary" value from cell
        cell_data_summary = get_cell_by_column_name(source_row, "Summary", column_map)
        if cell_data_summary == None or cell_data_summary == 0:
            return None, 0
        summary = cell_data_summary.display_value

        # make sure the row matches the primary and description and chapter
        if (summary == jira_entry[2]) and (description == jira_entry[3]) and (chapter == jira_entry[16]) and (primary == jira_entry[18]):
            # Build new cell value
            new_cell = ss_client.models.Cell()
            new_cell.column_id = column_map["Input into JIRA"]
            new_cell.value = True
            new_cell.strict = False
            # Build the row to update
            new_row = ss_client.models.Row()
            new_row.id = source_row.id
            new_row.cells.append(new_cell)
            return new_row, source_row.id
    return None, 0


def updateSmartsheetMain(program_input_path, jira_input_path):
    ''' walk through file that was imported into Jira
    and update the imported status in smartsheets to true'''
    colorama.init()
    #program_input_path = sys.argv[1]
    #jira_input_path = sys.argv[2]
    print("Python Version is", platform.python_version())
    print("System Version is", platform.platform())
    # The API identifies columns by Id,
    # but it's more convenient to refer to column names. Store a map here
    column_map = {}
    #keep track of how many update matches we find
    matches = 0
    print("")
    access_token = get_access_token(program_input_path)
    # Initialize client
    ss_client = smartsheet.Smartsheet(access_token)
    # Make sure we don't miss any error
    ss_client.errors_as_exceptions(True)

    # Log all calls when debugging
    #logging.basicConfig(filename=program_input_path + 'rwsheet.log', filemode='w', level=logging.INFO)
    # The API identifies columns by Id,
    # but it's more convenient to refer to column names. Store a map here
    # get a list of Smartsheet IDs we will use to update sheets
    ids_list = get_smartsheet_ids(program_input_path)
    #let's get the data from the JIRA import file so we can update the Smartsheets
    jira_list = read_jira_import_file(jira_input_path)
    # create a dictionary so we can keep track of of the JIRA entries we found a match for
    jira_dict = {}
    #we only want to create this above dictionary the first time thru but not every interation thru smartsheet sheets
    jira_list_first_time = 1
    #we are going to walk thru all entries in the JIRA import file
    for sheet_id in ids_list:
        # set the next 2 values so that we skip the tile row on the JIRA import file
        jira_entry_count = 1
        jira_dict[0] = 'true'
        #get the sheet info from Smartsheets
        sheet = ss_client.Sheets.get_sheet(sheet_id)
        print("Checking Smartsheet", sheet.name)
        # Build column map for later reference - translates column names to column id
        for column in sheet.columns:
            column_map[column.title] = column.id
        # Accumulate rows needing update here. we will do a bulk update to Smartsheeets
        rows_to_update = []
        #keep track of row ids to make sure there's no duplicate for call to smartsheets which will cause an error
        row_ids = []
        #walk the list of JIRA entries
        for jira_entry in jira_list:
            #dont; try to lookup first row since it only contains the column headers
            if jira_entry == jira_list[0]:
                continue
            #if it's the first time thru the JIRA list let setup the dictionary to false until we find a match
            if jira_list_first_time == 1:
                jira_dict[jira_entry_count] = 'false'
            #walk all the rows in a sheet
            for row in sheet.rows:
                row_to_update, row_id = evaluate_row_and_build_updates(ss_client, row, jira_entry, column_map)
                # we found a match then build a list of rows that we will eventually upload to Smartsheets
                if row_to_update != None:
                    if row_id not in row_ids:
                        rows_to_update.append(row_to_update)
                        row_ids.append(row_id)
                    else:
                        print(Bcolors.FAIL + "Duplicate found in JIRA import file row " + str(jira_entry_count + 1) + " in Smartsheet " + "'" + sheet.name + "'" + " row " + str(row.row_number) + Bcolors.ENDC)
                    matches += 1
                    jira_dict[jira_entry_count] = 'true'
                    print(Bcolors.OKGREEN + "Found JIRA import file row " + str(jira_entry_count + 1) + " in Smartsheet " + "'" + sheet.name + "'" + " row " + str(
                        row.row_number) + Bcolors.ENDC)
                    break
            jira_entry_count += 1
        # Finally, write updated cells back to Smartsheet in bulk
        retry = 3
        while retry != 0:
            if rows_to_update:
                result = ss_client.Sheets.update_rows(sheet_id, rows_to_update)
                #setup to do retries in case we hit the rate limit
                if result.result_code != 0:
                    print(Bcolors.FAIL + "Error updating Smartsheet " + sheet.name + "delaying for 1 minute before retry" + Bcolors.ENDC )
                    retry = retry - 1
                    time.sleep(60)
                else:
                    print()
                    retry = 0
            else:
                print("No updates required in ", sheet.name)
                print()
                retry = 0
        jira_list_first_time = 0
    dict_count = 0
    #now that we are done, print out all row that are false indicating they were not updated
    for dict_index, dict_entry in jira_dict.items():
        if jira_dict[dict_index] == 'false':
            dict_count += 1
            print(Bcolors.WARNING + "JIRA import file row " + str(dict_index+1) + " not found or already checked in Smartsheets" + Bcolors.ENDC)

    print(Bcolors.OKGREEN + "\nTotal Smartsheet Rows not Updated = " + str(dict_count) + Bcolors.ENDC)
    print(Bcolors.OKGREEN + "Total Smartsheet Rows Updated     = " + str(matches) + Bcolors.ENDC)
    colorama.deinit()

if __name__ == "__main__":
    updateSmartsheetMain(program_input_path, jira_input_path)
