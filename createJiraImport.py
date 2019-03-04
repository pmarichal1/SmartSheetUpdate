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
import logging
import time

import colorama


class Bcolors():
    '''color pallette for print statements'''
    HEADER = '\033[96m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'


def read_file_info(filename):
    '''create a list from the file information'''
    filelist = []
    try:
        with open(filename, mode='r') as file_hdnl:
            reader = csv.reader(file_hdnl, delimiter='\t')
            for line in reader:
                filelist.append(line)
            file_hdnl.close()
            print(Bcolors.HEADER + f"{len(filelist)} - Number of entries found in {filename}" + Bcolors.ENDC)
            return filelist
    except IOError:
        print(Bcolors.FAIL + "No such file:" + filename + Bcolors.ENDC)
        sys.exit(1)


def create_primarymod_suffix (jira_list, jira_map):
    ''' need to create a suffix version of Primary so we can looked it yp against QJT entries'''
    for jira_entry in jira_list:
    # need to create a list member by taking sheet name and adding siffux to primary
        jira_index = get_index_by_column_name("sheet name", jira_map)
        sheetname=jira_entry[jira_index]
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
    '''Helper function to find index from column name'''
    try:
        # make sure the key we are looking up exists
        column_index = column_map.get(column_name)
        if column_index is None:
            print(Bcolors.FAIL + f"Column name '{column_name}' returns None" + Bcolors.ENDC)
        return column_index
    except KeyError:
        print(Bcolors.FAIL + f"Column name '{column_name}' does not exist" + Bcolors.ENDC)
        return 0


def print_list_entry (mylist, column_name, mymap):
    ''' print the lsit'''
    for entry in mylist:
        index = get_index_by_column_name(column_name, mymap)
        if index is not None:
            print (f" Entry {entry[index]}")
    return 0


def create_final_list(jira_list, jira_map, qjt_list, processing_errors):
    ''' looking for Primary field matches in QJT to create final list of valid import entries'''
    notfount = 0
    entries_found = 0
    passes = 2
    print(Bcolors.HEADER + f" ********** Seaching for Primary matches **********" + Bcolors.ENDC)
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
                    #print("FOUNT IT without suffix")
                    foundit = 1
                    entries_found += 1
                    break
        # could not find it at all
        if foundit != 1:
            print(Bcolors.WARNING + "Primary {} line {:3} of all JIRA Bugs.txt not found".format(entry_jira[get_index_by_column_name('primary', jira_map)], passes)+ Bcolors.ENDC)
            processing_errors.append("Primary {} line {:3} of all JIRA Bugs.txt not found".format(entry_jira[get_index_by_column_name('primary', jira_map)], passes))
            #processing_errors.append(f"Primary {entry_jira[get_index_by_column_name('primary', jira_map)]} line {passes} of all JIRA Bugs.txt not found")
            notfount += 1
        passes += 1
    print(Bcolors.FAIL + f"{notfount+1} - Total Primary entries not found" + Bcolors.ENDC)
    print(Bcolors.OKGREEN + f"{entries_found+1} - Total Primary entries found" + Bcolors.ENDC)
    return import_list, processing_errors

def cleanup_import_list(my_list):
    ''' re-order the final list '''
    ret_list = []
    myorder = [4,3,10,8,1,2,11,9,5,6,7]
    for entry in my_list:
        del entry[0]
        del entry[6:9]
        del entry[11]
        #reorder the list
        values_by_index = dict(zip(myorder, entry))
        new_entry = [values_by_index[index + 1] for index in range(len(myorder))]
        ret_list.append(new_entry)
    return ret_list


def fill_in_static_columns(final_import_list):
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


def lookup_V42_username(final_import_list, final_map, username_list, username_map, processing_errors):
    ''' lookup and convert text string of username into V42 version '''
    print(Bcolors.HEADER + f" ********** Lookup V42 entries **********" + Bcolors.ENDC)
    # start with line 2 since we are adding a header at the end of function
    notfound = 0
    passes = 2
    for entry in final_import_list:
        found = 0
        # save original username in case we need to print it out
        reporter = entry[get_index_by_column_name('Reporter', final_map)]
        for entry1 in username_list:
            if entry[get_index_by_column_name('Reporter', final_map)] in entry1:
                entry[get_index_by_column_name('Reporter', final_map)] = entry1[get_index_by_column_name('V42name', username_map)]
                found = 1
        if found != 1:
            entry[get_index_by_column_name('Reporter', final_map)] = username_list[0][get_index_by_column_name('V42name', username_map)]
            print(Bcolors.WARNING + "Username '{:>20}' line {:>3} of Final Output file not found. Mapping to {}"
                .format(reporter, passes, username_list[0][get_index_by_column_name('V42name', username_map)]) + Bcolors.ENDC)
            processing_errors.append("Username '{:>20}' line {:>3} of Final Output file not found. Mapping to {}"
                .format(reporter, passes, username_list[0][get_index_by_column_name('V42name', username_map)]))
            notfound += 1

        passes += 1
    final_import_list.insert(0, ["Project", "Issue Type", "Summary",
                                 "Description", "Epic Link", "Severity", "Reporter", "Found in Environment",
                                 "Found in Phase", "Labels", "Priority", "Activity", "Custom1", "Custom2",
                                 "Custom3", "Custom4", "Custom5", "Edition", "Primary", "Title"])
    print(Bcolors.FAIL + f"{notfound} -  Total of Usernames not matched" + Bcolors.ENDC)
    return final_import_list


def fill_in_epic_link(my_list, final_map, epic_list, epic_map, processing_errors):
    notfound = 0
    entries_found = 0
    passes = 2
    final_import_list = []
    print(Bcolors.HEADER + f" ********** Seaching for Epic Link matches **********" + Bcolors.ENDC)
    for entry in my_list:
        found = 0
        for entry1 in epic_list:
            dstr = entry[get_index_by_column_name('Custom3', final_map)]
            #check is string is empty
            if not dstr:
                break
            if dstr in entry1[get_index_by_column_name('custom', epic_map)]:
                entry[get_index_by_column_name('Epic Link', final_map)] = entry1[get_index_by_column_name('key', epic_map)]
                final_import_list.append(entry)
                entries_found += 1
                found = 1
                break
        if found == 0:
            if not dstr:
                notfound += 1
                continue
            for entry1 in epic_list:
                if dstr in entry1[get_index_by_column_name('summary', epic_map)]:
                    entry[get_index_by_column_name('Epic Link', final_map)] = entry1[get_index_by_column_name('key', epic_map)]
                    final_import_list.append(entry)
                    found = 1
                    entries_found += 1
                    break
        if found == 0:
            print(Bcolors.WARNING + "DemandID {:>7} associated with Primary {} not found".format(entry[get_index_by_column_name('Custom3', final_map)], entry[get_index_by_column_name('Primary', final_map)]) + Bcolors.ENDC)
            processing_errors.append("DemandID {:>7} associated with Primary {} not found".format(entry[get_index_by_column_name('Custom3', final_map)], entry[get_index_by_column_name('Primary', final_map)]))
            notfound += 1
        passes +=1
    print(Bcolors.FAIL + f"{notfound} -  Total of DemandIDs not found" + Bcolors.ENDC)
    print(Bcolors.OKGREEN + f"{entries_found} -  Total of DemandIDs found" + Bcolors.ENDC)
    return final_import_list


def write_list_to_file(filename, import_list):
    with open(filename, 'w') as file_hndl:
        try:
            file_hndl.writelines('\t'.join(i) + '\n' for i in import_list)
        except IOError:
            print(Bcolors.FAIL + "No such file:" + filename + Bcolors.ENDC)
            sys.exit(1)
    #output list to a file

    return(import_list)

def write_errors_to_file(filename, error_list):
    with open(filename, 'w') as file_hndl:
        try:
            for line in error_list:
                file_hndl.write(str(line))
                file_hndl.write('\n')
        except IOError:
            print(Bcolors.FAIL + "No such file:" + filename + Bcolors.ENDC)
            sys.exit(1)
    #output list to a file

    return(error_list)

def create_jira_import_main(program_input_path, jira_input_path):
    colorama.init()
    print("Python Version is", platform.python_version())
    print("System Version is", platform.platform())
    localtime = time.asctime(time.localtime(time.time()))
    print("Local current time :", localtime)
    print("")

    # store text to index maps here
    jira_map = {"sheet name": 0,"editorial": 1,"freelancer": 2, "primary":3, "chapter": 4,"summary": 5,"additional": 6, "primarymod": 7}
    qjt_map = {"primary": 0,"assingnedpm": 1,"title": 2,"edition": 3,"isbn": 4,"demandid":5,"author": 6}
    epic_map = {"type": 0,"key": 1,"id": 2,"summary": 3,"custom": 4}
    final_map = {"Project": 0 , "Issue Type": 1, "Summary" : 2,"Description": 3, "Epic Link": 4, "Severity": 5, "Reporter": 6,
                                "Found in Environment": 7, "Found in Phase": 8, "Labels": 9, "Priority": 10, "Activity": 11,
                                "Custom1": 12, "Custom2": 13, "Custom3": 14, "Custom4": 15, "Custom5": 16, "Edition": 17, "Primary": 18, "Title": 19}
    username_map = {"name": 0,"V42name": 1}

    # Log all calls when debugging
    logging.basicConfig(filename=program_input_path + 'rwsheet.log', filemode='w', level=logging.INFO)
    processing_errors = []
    processing_errors.append("Created on " + localtime)
    # The API identifies columns by Id,
    # but it's more convenient to refer to column names. Store a map here
    # get a list of Smartsheet IDs we will use to update sheets
    print(Bcolors.HEADER + f" ********** Gathering file information **********" + Bcolors.ENDC)
    filename =  program_input_path + 'All JIRA bugs.txt'
    jira_list = read_file_info(filename)
    create_primarymod_suffix (jira_list, jira_map)

    filename = program_input_path + 'Quality Job Tracker for JIRA.txt'
    qjt_list = read_file_info(filename)
    #print_list_entry(qjt_list, "isbn", qjt_map)

    filename = program_input_path + "EpicsQUA.txt"
    epic_list = read_file_info(filename)
    filename = program_input_path + "EpicsPM.txt"
    epic_list += read_file_info(filename)
    filename = program_input_path + "EpicsPRODN.txt"
    epic_list += read_file_info(filename)

    filename = program_input_path + "UserNames.txt"
    username_list = read_file_info(filename)
    #print_list_entry(username_list, "V42name", username_map)

    final_import_list, processing_errors = create_final_list(jira_list, jira_map, qjt_list, processing_errors)
    final_import_list = cleanup_import_list(final_import_list)
    final_import_list = fill_in_static_columns(final_import_list)
    final_import_list = fill_in_epic_link(final_import_list, final_map, epic_list ,epic_map, processing_errors)
    final_import_list = lookup_V42_username(final_import_list, final_map, username_list, username_map, processing_errors)
    filename = jira_input_path + 'JIRAImportData.txt'
    write_list_to_file(filename, final_import_list)
    filename =  jira_input_path + 'ProcessingErrors.txt'
    write_errors_to_file(filename, processing_errors)
    print(Bcolors.OKGREEN + f"{len(final_import_list)-1} - Total entries written to output file" + Bcolors.ENDC)
    colorama.deinit()

if __name__ == "__main__":
    program_input_path = sys.argv[1]
    jira_input_path = sys.argv[2]
    create_jira_import_main(program_input_path, jira_input_path)
