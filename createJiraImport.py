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


def read_file_info(filename, header):
    """create a list from the file information. header=0  means no header in file,header=1 means header present"""
    filelist = []
    try:
        with open(filename, mode='r') as file_hdnl:
            reader = csv.reader(file_hdnl, delimiter='\t')
            for line in reader:
                filelist.append(line)
            file_hdnl.close()
            print(Bcolors.HEADER + "{:>3} - Number of entries found in {}".format(len(filelist) - header,
                                                                                  filename) + Bcolors.ENDC)
            return filelist
    except IOError:
        print(Bcolors.FAIL + "No such file:" + filename + Bcolors.ENDC)
        sys.exit(1)


def create_primarymod_suffix(jira_list, jira_map):
    """ need to create a suffix version of Primary so we can looked it yp against QJT entries"""
    for jira_entry in jira_list:
        # need to create a list member by taking sheet name and adding siffux to primary
        jira_index = get_index_by_column_name("sheet name", jira_map)
        sheetname = jira_entry[jira_index]
        jira_index = get_index_by_column_name("primarymod", jira_map)
        if sheetname.find('Copyediting') != -1:
            primarymod = (jira_entry[get_index_by_column_name("primary", jira_map)]) + "-C"
            jira_entry[jira_index] = primarymod
        elif sheetname.find('Proofreading') != -1:
            primarymod = (jira_entry[get_index_by_column_name("primary", jira_map)]) + "-P"
            jira_entry[jira_index] = primarymod
        elif sheetname.find('Indexing') != -1:
            primarymod = (jira_entry[get_index_by_column_name("primary", jira_map)]) + "-I"
            jira_entry[jira_index] = primarymod
    return 0


def get_index_by_column_name(column_name, column_map):
    """Helper function to find index from column name"""
    try:
        # make sure the key we are looking up exists
        column_index = column_map.get(column_name)
        if column_index is None:
            print(Bcolors.FAIL + f"Column name '{column_name}' returns None" + Bcolors.ENDC)
        return column_index
    except KeyError:
        print(Bcolors.FAIL + f"Column name '{column_name}' does not exist" + Bcolors.ENDC)
        return 0


def print_list_entry(mylist, column_name, mymap):
    """ print the lsit"""
    for entry in mylist:
        index = get_index_by_column_name(column_name, mymap)
        if index is not None:
            print(f" Entry {entry[index]}")
    return 0


def lookup_primary_entry(jira_list, jira_map, qjt_list, processing_errors):
    """ looking for Primary field matches in QJT to create final list of valid import entries"""
    not_found = 0
    entries_found = 0
    row_number = 2
    print(Bcolors.HEADER + f" ********** Seaching for Job # matches **********" + Bcolors.ENDC)
    # don't need the first row since its the headers
    del jira_list[0]
    import_list = []

    for entry_jira in jira_list:
        foundit = 0
        # look for match with suffix
        for entry_qjt in qjt_list:
            if entry_jira[get_index_by_column_name('primarymod', jira_map)] in entry_qjt:
                import_list.append(entry_jira + entry_qjt)
                foundit = 1
                entries_found += 1
                break
        # see if it exists without suffix
        if foundit != 1:
            for entry_qjt in qjt_list:
                if entry_jira[get_index_by_column_name('primary', jira_map)] in entry_qjt:
                    import_list.append(entry_jira + entry_qjt)
                    # print("FOUNT IT without suffix")
                    foundit = 1
                    entries_found += 1
                    break
        # could not find it at all
        if foundit != 1:
            print(
                Bcolors.WARNING + "Job # '{}' line '{:3}' of 'All JIRA Bugs.txt' not found in 'Quality Job Tracker for JIRA.txt' "
                .format(entry_jira[get_index_by_column_name('primary', jira_map)], row_number) + Bcolors.ENDC)
            processing_errors.append(
                "Job # '{}' line '{:3}' of 'All JIRA Bugs.txt' not found in 'Quality Job Tracker for JIRA.txt'"
                    .format(entry_jira[get_index_by_column_name('primary', jira_map)], row_number))
            # processing_errors.append(f"Primary {entry_jira[get_index_by_column_name('primary', jira_map)]} line {passes} of all JIRA Bugs.txt not found")
            not_found += 1
        row_number += 1
    print(Bcolors.FAIL + f"{not_found} - Total 'Job #' entries not found" + Bcolors.ENDC)
    print(Bcolors.OKGREEN + f"{entries_found} - Total 'Job #'' entries found" + Bcolors.ENDC)
    return import_list


def cleanup_import_list(my_list):
    """ re-order the final list """
    ret_list = []
    myorder = [4, 3, 10, 8, 1, 2, 11, 9, 5, 6, 7]
    for entry in my_list:
        del entry[0]
        del entry[6:9]
        del entry[11]
        # reorder the list
        values_by_index = dict(zip(myorder, entry))
        new_entry = [values_by_index[index + 1] for index in range(len(myorder))]
        ret_list.append(new_entry)
    return ret_list


def fill_in_static_columns(final_import_list):
    """ add static info to cloumns"""
    for entry in final_import_list:
        entry.insert(0, "CSC_Pearson_Quality")
        entry.insert(1, "Bug")
        entry.insert(5, "Medium")
        entry.insert(7, "QA")
        entry.insert(8, "Production")
        entry.insert(9, "L3")
        entry.insert(10, "Minor")
        entry.insert(11, "QA")
        entry.insert(12, "SPi")
    return final_import_list


def lookup_v42_username(final_import_list, final_map, username_list, username_map, processing_errors):
    """ lookup and convert text string of username into V42 version """
    print(Bcolors.HEADER + f" ********** Lookup V42 entries **********" + Bcolors.ENDC)
    # start with line 2 since we are adding a header at the end of function
    not_found = 0
    row_number = 2
    for entry in final_import_list:
        foundit = 0
        # save original username in case we need to print it out
        reporter = entry[get_index_by_column_name('Reporter', final_map)]
        for entry1 in username_list:
            if entry[get_index_by_column_name('Reporter', final_map)] in entry1:
                entry[get_index_by_column_name('Reporter', final_map)] = entry1[
                    get_index_by_column_name('V42name', username_map)]
                foundit = 1
        if foundit != 1:
            entry[get_index_by_column_name('Reporter', final_map)] = username_list[0][
                get_index_by_column_name('V42name', username_map)]
            print(Bcolors.WARNING + "Username '{:>20}' line '{:>3}' of JIRAImportData file not found in 'Usernames.txt'. Mapping to '{}'"
                .format(reporter, row_number, username_list[0][get_index_by_column_name('V42name', username_map)]) + Bcolors.ENDC)
            processing_errors.append("Username '{:>20}' line '{:>3}' of JIRAImportData file not found in 'Usernames.txt. Mapping to '{}'"
                    .format(reporter, row_number, username_list[0][get_index_by_column_name('V42name', username_map)]))
            not_found += 1

        row_number += 1
    final_import_list.insert(0, ["Project", "Issue Type", "Summary", "Description", "Epic Link", "Severity",
                                 "Reporter", "Found in Environment", "Found in Phase", "Labels", "Priority",
                                 "Activity", "Custom1", "Custom2", "Custom3", "Custom4", "Custom5", "Edition",
                                 "Job #", "Title"])
    print(Bcolors.FAIL + f"{not_found} -  Total of Usernames not matched" + Bcolors.ENDC)
    return final_import_list


def fill_in_epic_link(final_list, final_map, epic_list, epic_map, processing_errors):
    """ lookup ISBN and DemandiD to see if there is a match in the Epic files
        This function will take out the ISBN from final output file and look thru all Epic files in this order
        If the  ISBN is represented by 2 numbers, split them and process each of them
        1st - look fir ISBN in ISBN column of epic files.
        Next - look for an ISBN  match in Summary column
        Next - look for a DemandID match of DemandId column
        Next - look for a DemandID match of Summary column """

    not_found = 0
    entries_found = 0
    row_number = 2
    final_import_list = []
    print(Bcolors.HEADER + f" ********** Searching for Epic Link matches **********" + Bcolors.ENDC)
    for entry in final_list:
        foundit = 0
        for entry1 in epic_list:
            #look for ISBN match first in ISBN column in Epic files
            datastr = entry[get_index_by_column_name('Custom2', final_map)]
            datastr = datastr.strip()
            # ISBN might be represented with 2 numbers.. split them and process both
            if '/' in datastr:
                isbn1, isbn2 = datastr.split('/')
                isbn1 = isbn1.strip()
                isbn2 = isbn2.strip()
                if isbn2:
                    if isbn2 in entry1[get_index_by_column_name('isbn', epic_map)]:
                        entry[get_index_by_column_name('Epic Link', final_map)] = entry1[
                            get_index_by_column_name('key', epic_map)]
                        final_import_list.append(entry)
                        isbn += 1
                        entries_found += 1
                        foundit = 1
                        break
                    # look for ISBN match first in Summary column in Epic files
                    if isbn2 in entry1[get_index_by_column_name('summary', epic_map)]:
                        entry[get_index_by_column_name('Epic Link', final_map)] = entry1[
                            get_index_by_column_name('key', epic_map)]
                        final_import_list.append(entry)
                        isbns += 1
                        foundit = 1
                        entries_found += 1
                        break
                datastr = isbn1
            # check is string is empty
            if datastr:
                if datastr in entry1[get_index_by_column_name('isbn', epic_map)]:
                    entry[get_index_by_column_name('Epic Link', final_map)] = entry1[
                        get_index_by_column_name('key', epic_map)]
                    final_import_list.append(entry)
                    entries_found += 1
                    foundit = 1
                    break
                #look for ISBN match first in Sumamry column in Epic files
                if datastr in entry1[get_index_by_column_name('summary', epic_map)]:
                    entry[get_index_by_column_name('Epic Link', final_map)] = entry1[
                        get_index_by_column_name('key', epic_map)]
                    final_import_list.append(entry)
                    foundit = 1
                    entries_found += 1
                    break
            # did not find ISBN match so let's see if we can match DemandID
            #look for DemandId match first in DeamndId column in Epic files
            datastr = entry[get_index_by_column_name('Custom3', final_map)]
            datastr = datastr.strip()
            if datastr:
                if datastr in entry1[get_index_by_column_name('demandid', epic_map)]:
                    entry[get_index_by_column_name('Epic Link', final_map)] = entry1[
                        get_index_by_column_name('key', epic_map)]
                    final_import_list.append(entry)
                    entries_found += 1
                    foundit = 1
                    break
                #look for DemandId match first in Sumamry column in Epic files
                if datastr in entry1[get_index_by_column_name('summary', epic_map)]:
                    entry[get_index_by_column_name('Epic Link', final_map)] = entry1[
                        get_index_by_column_name('key', epic_map)]
                    final_import_list.append(entry)
                    foundit = 1
                    entries_found += 1
                    break
        if foundit == 0:
            print(Bcolors.WARNING + "No Epic match '{:>7}' associated with Job # '{}' found in Epic files ".format(
                entry[get_index_by_column_name('Custom3', final_map)],
                entry[get_index_by_column_name('Job #', final_map)]) + Bcolors.ENDC)
            processing_errors.append("No Epic match '{:>7}' associated with Job # '{}' found in Epic files ".format(
                entry[get_index_by_column_name('Custom3', final_map)],
                entry[get_index_by_column_name('Job #', final_map)]))
            not_found += 1
        row_number += 1
    print(Bcolors.FAIL + f"{not_found} -  Total of Epics not found" + Bcolors.ENDC)
    print(Bcolors.OKGREEN + f"{entries_found} -  Total of Epics found" + Bcolors.ENDC)
    return final_import_list


def write_list_to_file(filename, import_list):
    """  write out final list to file """
    try:
        with open(filename, 'w') as file_hndl:
            try:
                file_hndl.writelines('\t'.join(i) + '\n' for i in import_list)
            except IOError:
                print(Bcolors.FAIL + "Cannot write to file: " + filename + Bcolors.ENDC)
                sys.exit(1)
    except FileNotFoundError:
        print(Bcolors.FAIL + "Cannot open file:" + filename + Bcolors.ENDC)
        sys.exit(1)


def write_errors_to_file(filename, error_list):
    """ write out items not resolved  """
    try:
        with open(filename, 'w') as file_hndl:
            try:
                for line in error_list:
                    file_hndl.write(str(line))
                    file_hndl.write('\n')
            except IOError:
                print(Bcolors.FAIL + "Cannot write to file: " + filename + Bcolors.ENDC)
                sys.exit(1)
    except FileNotFoundError:
        print(Bcolors.FAIL + "Cannot open file:" + filename + Bcolors.ENDC)
        sys.exit(1)


def validate_final_import_file(final_import_list, final_map):
    """ look through final output to see if there are any empty entries"""
    print(Bcolors.HEADER + f" ********** Validating  JIRAImportData **********" + Bcolors.ENDC)
    row_number = 1
    revdict = dict([[v, k] for k, v in final_map.items()])
    empty_list = []
    for line in final_import_list:
        empty_list = []
        for empty_value in range(len(final_map)):
            if line[empty_value] == '':
                empty_list.append(revdict[empty_value])
        if empty_list:
            blank_line = (', '.join(empty_list))
            print(Bcolors.WARNING + "Line '{:>3}' of JIRAImportData file has empty columns '{}'".format(row_number,
                                                                                                        blank_line) + Bcolors.ENDC)
        row_number += 1


def create_filename_dict(filename_list):
    filename_dict = {}
    filename_index = 0
    for line in filename_list:
        full_str = ''.join(line)
        filename_dict[filename_index] = full_str
        filename_index += 1
    return filename_dict


def create_jira_import_main(program_input_path, jira_input_path):
    """ main code """
    colorama.init()
    localtime = time.asctime(time.localtime(time.time()))
    # store text to index maps here
    jira_map = {"sheet name": 0, "editorial": 1, "freelancer": 2, "primary": 3, "chapter": 4, "summary": 5,
                "additional": 6, "primarymod": 7}
    qjt_map = {"primary": 0, "assingnedpm": 1, "title": 2, "edition": 3, "isbn": 4, "demandid": 5, "author": 6}
    epic_map = {"key": 0, "summary": 1, "demandid": 2, "isbn": 3}
    final_map = {"Project": 0, "Issue Type": 1, "Summary": 2, "Description": 3, "Epic Link": 4, "Severity": 5,
                 "Reporter": 6, "Found in Environment": 7, "Found in Phase": 8, "Labels": 9, "Priority": 10,
                 "Activity": 11, "Custom1": 12, "Custom2": 13, "Custom3": 14, "Custom4": 15, "Custom5": 16,
                 "Edition": 17, "Job #": 18, "Title": 19}
    username_map = {"name": 0, "V42name": 1}

    # Log all calls when debugging
    # logging.basicConfig(filename=program_input_path + 'rwsheet.log', filemode='w', level=logging.INFO)
    processing_errors = []
    processing_errors.append("Created on " + localtime)
    # The API identifies columns by Id,
    # but it's more convenient to refer to column names. Store a map here
    # get a list of Smartsheet IDs we will use to update sheets
    print(Bcolors.HEADER + f" ********** Gathering file information **********" + Bcolors.ENDC)
    filename = program_input_path + 'filenames.txt'
    filename_list = read_file_info(filename, 0)
    filename_dict = create_filename_dict(filename_list)
    # get JIRA info
    if len(filename_list) < 6:
        print(Bcolors.FAIL + "Not enough entries in filename.txt , should be 6 entries" + Bcolors.ENDC)
        sys.exit(1)
    # get 'All JIRA bugs.txt' file info
    filename = program_input_path + filename_dict.get(0, "")
    jira_list = read_file_info(filename, 1)
    create_primarymod_suffix(jira_list, jira_map)
    # get Quality Job Tracker info
    filename = program_input_path + filename_dict.get(1, "")
    qjt_list = read_file_info(filename, 1)
    # get EpicsQUA info
    filename = program_input_path + filename_dict.get(2, "")
    epic_list = read_file_info(filename, 1)
    # get EpicsPM info
    filename = program_input_path + filename_dict.get(3, "")
    epic_list += read_file_info(filename, 1)
    # get EpicsPRODN info
    filename = program_input_path + filename_dict.get(4, "")
    epic_list += read_file_info(filename, 1)
    # get 'UserNames.txt'  info
    filename = program_input_path + filename_dict.get(5, "")
    username_list = read_file_info(filename, 0)
    if len(username_list) < 1:
        print(Bcolors.FAIL + "Need at least 1 entry in UserNames.txt" + Bcolors.ENDC)
        sys.exit(1)
    final_import_list = lookup_primary_entry(jira_list, jira_map, qjt_list, processing_errors)
    final_import_list = cleanup_import_list(final_import_list)
    final_import_list = fill_in_static_columns(final_import_list)
    final_import_list = fill_in_epic_link(final_import_list, final_map, epic_list, epic_map, processing_errors)
    final_import_list = lookup_v42_username(final_import_list, final_map, username_list, username_map,
                                            processing_errors)
    validate_final_import_file(final_import_list, final_map)
    print(Bcolors.HEADER + f" ********** Writing  JIRAImportData **********" + Bcolors.ENDC)
    print(Bcolors.OKGREEN + f"{len(final_import_list) - 1} - Total entries written to output file" + Bcolors.ENDC)
    filename = jira_input_path + 'JIRAImportData.txt'
    write_list_to_file(filename, final_import_list)
    filename = jira_input_path + 'ProcessingErrors.txt'
    write_errors_to_file(filename, processing_errors)
    colorama.deinit()


