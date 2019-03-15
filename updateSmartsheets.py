"""
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
"""
import csv
import platform
import sys
# import logging
import time
import colorama
import smartsheet
from smartsheet import sheets


class Bcolors:
    """color pallette for print statements"""
    HEADER = '\033[96m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'


def get_file_info(filename):
    """# go get the Smartsheet IDs we want to download"""
    try:
        ss_hndl = open(filename, 'r')
    except FileNotFoundError:
        print(Bcolors.FAIL + f"File not found {filename}" + Bcolors.ENDC)
        sys.exit(1)
    except Exception as e:
        errno, strerror = e.args
        print(Bcolors.FAIL + f"I/O error({errno}): {strerror} {filename}" + Bcolors.ENDC)
        sys.exit(1)
    else:
        line_list = ss_hndl.readlines()
        # strip off newlines
        info_list = []
        for line in line_list:
            info_list.append(line.rstrip())
        ss_hndl.close()
        return info_list


def read_jira_import_file(filename):
    """create a tuple list from the JIRAInputData file"""
    try:
        jira_hdnl = open(filename, mode='r')
    except FileNotFoundError:
        print(Bcolors.FAIL + f"File not found {filename}" + Bcolors.ENDC)
        sys.exit(1)
    except Exception as e:
        errno, strerror = e.args
        print(Bcolors.FAIL + f"I/O error({errno}): {strerror} {filename}" + Bcolors.ENDC)
        sys.exit(1)
    else:
        reader = csv.reader(jira_hdnl, delimiter='\t')
        jira_list = list(map(tuple, reader))
        jira_hdnl.close()
        return jira_list


def get_cell_by_column_name(row_id, column_name, column_map):
    """Helper function to find cell in a row and return name"""
    try:
        # make sure the key we are looking up exists
        column_id = column_map[column_name]
        if row_id.get_column(column_id) is None:
            print(Bcolors.FAIL + f"Column name '{column_name}' returns None" + Bcolors.ENDC)
        return row_id.get_column(column_id)
    except KeyError:
        print(Bcolors.FAIL + f"Column name '{column_name}' does not exist" + Bcolors.ENDC)
        return 0


def evaluate_row_and_build_updates(ss_client, source_row, jira_entry, column_map, empty_dict, sheet_name):
    """Find the cell and value we want to evaluate"""
    #   for "Input into JIRA" value from the cell
    cell_data_checkbox = get_cell_by_column_name(source_row, "Input into JIRA", column_map)
    if (cell_data_checkbox is None) or (cell_data_checkbox == 0):
        return None, 0
    checkbox = cell_data_checkbox.value
    if checkbox is not True:  # Skip if already true
        # get contents of description column
        cell_data_description = get_cell_by_column_name(source_row, "Additional description if necessary", column_map)
        # get "Additional description if necessary" value from cell
        if cell_data_description is None or cell_data_description == 0:
            return None, 0
        description = cell_data_description.display_value

        # get "JOB #" value from cell
        cell_data_job = get_cell_by_column_name(source_row, "Job #", column_map)
        if cell_data_job is None or cell_data_job == 0:
            return None, 0
        primary = cell_data_job.display_value

        # get "chapter #s" value from cell
        cell_data_chapter = get_cell_by_column_name(source_row, "Chapter # or #s", column_map)
        if cell_data_chapter is None or cell_data_chapter == 0:
            return None, 0
        chapter = cell_data_chapter.display_value

        # get "Summary" value from cell
        cell_data_summary = get_cell_by_column_name(source_row, "Summary", column_map)
        if cell_data_summary is None or cell_data_summary == 0:
            return None, 0
        summary = cell_data_summary.display_value
        if summary != '[No Errors to Report]':
            if (summary is None) or (chapter is None) \
                    or (primary is None) or (description is None):
                if (source_row.row_number, sheet_name) not in empty_dict:
                    empty_dict[source_row.row_number, sheet_name] = 1
                else:
                    empty_dict[source_row.row_number, sheet_name] += 1
        # make sure the row matches the primary and description and chapter
        if (summary == jira_entry[2]) and (description == jira_entry[3]) \
                and (chapter == jira_entry[16]) and (primary == jira_entry[18]):
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

def write_updates_to_file(filename, update_list):
    """ write out items updated  """
    try:
        with open(filename, 'w') as file_hndl:
            try:
                for line in update_list:
                    file_hndl.write(str(line))
                    file_hndl.write('\n')
            except IOError:
                print(Bcolors.FAIL + "Cannot write to file: " + filename + Bcolors.ENDC)
                sys.exit(1)
    except FileNotFoundError:
        print(Bcolors.FAIL + "Cannot open file:" + filename + Bcolors.ENDC)
        sys.exit(1)


def updateSmartsheetMain(program_input_path, jira_input_path, user_credentials):
    """ walk through file that was imported into Jira
    and update the imported status in smartsheets to true"""
    colorama.init()
    # The API identifies columns by Id,
    # but it's more convenient to refer to column names. Store a map here
    column_map = {}
    # keep track of how many update matches we find
    matches = 0
    print("")
    # create a list to keep track of updates for output to file
    ss_update_list = []
    # Initialize client
    ss_client = smartsheet.Smartsheet(user_credentials[0])
    # Make sure we don't miss any error
    ss_client.errors_as_exceptions(True)

    # Log all calls when debugging
    # logging.basicConfig(filename=program_input_path + 'rwsheet.log', filemode='w', level=logging.INFO)
    # The API identifies columns by Id,
    # but it's more convenient to refer to column names. Store a map here
    # get a list of Smartsheet IDs we will use to update sheets
    filename = program_input_path + "smartsheetUpdateIDs.txt"
    ids_list = get_file_info(filename)
    if len(ids_list) < 3:
        print(Bcolors.FAIL + "Not enough entries in smartsheetUpdateIDs.txt , should be 3 entries" + Bcolors.ENDC)
        sys.exit(1)
    # let's get the data from the JIRA import file so we can update the Smartsheets
    filename = jira_input_path + 'JIRAImportData.txt'
    jira_list = read_jira_import_file(filename)
    # create a dictionary so we can keep track of of the JIRA entries we found a match for
    jira_dict = {}
    # keep track of row and sheetname for empty cells
    empty_cell = {}
    # only want to create this above dictionary the first time thru but not every iteration thru smartsheet sheets
    jira_list_first_time = 1
    # are going to walk thru all entries in the JIRA import file
    for sheet_id in ids_list:
        # set the next 2 values so that we skip the tile row on the JIRA import file
        jira_entry_count = 1
        jira_dict[0] = 'true'
        # get the sheet info from Smartsheets
        try:
            sheet = ss_client.Sheets.get_sheet(sheet_id)
        except Exception:
            print("Possible invalid sheet ID in smartsheetUpdateIDs.txt")
            sys.exit(1)
        print(Bcolors.OKGREEN + f"Checking Smartsheet {sheet.name}" + Bcolors.ENDC)
        # Build column map for later reference - translates column names to column id
        for column in sheet.columns:
            column_map[column.title] = column.id
        # Accumulate rows needing update here. we will do a bulk update to Smartsheeets
        rows_to_update = []
        # keep track of row ids to make sure there's no duplicate for call to smartsheets which will cause an error
        row_ids = []
        # walk the list of JIRA entries
        for jira_entry in jira_list:
            # don't try to lookup first row since it only contains the column headers
            if jira_entry == jira_list[0]:
                continue
            # if it's the first time thru the JIRA list let setup the dictionary to false until we find a match
            if jira_list_first_time == 1:
                jira_dict[jira_entry_count] = 'false'
            # walk all the rows in a sheet
            for row in sheet.rows:
                row_to_update, row_id = evaluate_row_and_build_updates(ss_client, row, jira_entry, column_map,
                                                                       empty_cell, sheet.name)
                # we found a match then build a list of rows that we will eventually upload to Smartsheets
                if row_to_update is not None:
                    if row_id not in row_ids:
                        rows_to_update.append(row_to_update)
                        row_ids.append(row_id)
                        matches += 1
                        ss_update_list.append("Found JIRA import file row " + str(
                            jira_entry_count + 1) + " in Smartsheet " + "'" + sheet.name + "'"
                              + " row " + str(row.row_number))
                        print(Bcolors.OKGREEN + "Found JIRA import file row " + str(
                            jira_entry_count + 1) + " in Smartsheet " + "'" + sheet.name + "'"
                              + " row " + str(row.row_number) + Bcolors.ENDC)
                    else:
                        ss_update_list.append("Duplicate found in JIRA import file row " + str(
                            jira_entry_count + 1) + " in Smartsheet " + "'" + sheet.name + "'" + " row " + str(
                            row.row_number))
                        print(Bcolors.FAIL + "Duplicate found in JIRA import file row " + str(
                            jira_entry_count + 1) + " in Smartsheet " + "'" + sheet.name + "'" + " row " + str(
                            row.row_number) + Bcolors.ENDC)
                    jira_dict[jira_entry_count] = 'true'
                    break
            jira_entry_count += 1
        # Finally, write updated cells back to Smartsheet in bulk
        # number of time to retry updating cells
        retry = 3
        # number of times we actually retried
        retries = 3
        while retry != 0:
            if rows_to_update:
                result = ss_client.Sheets.update_rows(sheet_id, rows_to_update)
                # setup to do retries in case we hit the rate limit
                if result.result_code != 0:
                    print(
                        Bcolors.FAIL + "Error updating Smartsheet " + sheet.name + "delaying for 1 minute before retry" + Bcolors.ENDC)
                    retry -= 1
                    retries -= 1
                    time.sleep(60)
                else:
                    print()
                    retry = 0
            else:
                print("No updates required in ", sheet.name)
                print()
                break
        jira_list_first_time = 0
        # if 0 then we failed to update
        if retries == 0:
            print(Bcolors.FAIL + "Failed to update Smartsheet " + sheet.name + " after 3 retries" + Bcolors.ENDC)
    dict_count = 0
    # now that we are done, print out all row that are false indicating they were not updated
    print(30 * "-" + Bcolors.OKBLUE + "Reporting Cells not Updated" + Bcolors.ENDC + 30 * "-")
    for dict_index, dict_entry in jira_dict.items():
        if jira_dict[dict_index] == 'false':
            dict_count += 1
            print(Bcolors.WARNING + "JIRA import file row " + "'" +
                  str(dict_index + 1) + "'" + " not found or already checked in Smartsheets" + Bcolors.ENDC)
    print(30 * "-" + Bcolors.OKBLUE + "Reporting Empty Cells" + Bcolors.ENDC + 30 * "-")
    for key, val in empty_cell.items():
        print(Bcolors.WARNING + f"Smartsheet '{key[1]}' empty cell in row '{key[0]}'" + Bcolors.ENDC)
    print(30 * "-" + Bcolors.OKBLUE + "Finished Updating Smartsheets" + Bcolors.ENDC + 30 * "-")
    print(Bcolors.OKGREEN + "\nTotal Smartsheet Rows not Updated = " + str(dict_count) + Bcolors.ENDC)
    print(Bcolors.OKGREEN + "Total Smartsheet Rows Updated     = " + str(matches) + Bcolors.ENDC)
    filename = jira_input_path + 'SmartsheetUpdates.txt'
    write_updates_to_file(filename, ss_update_list)
    colorama.deinit()


