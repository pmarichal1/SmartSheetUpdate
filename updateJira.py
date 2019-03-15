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

reference dict with JIRA issue
    issue_dict ={
            'project': 'QUA',                               # fixed
            'issuetype': {'name': 'Bug'},                   # fixed
            'customfield_10080' : {'value' : 'Medium'},     # fixed - Severity
            #'reporter' : 'VGagnDa',
            'customfield_10061' : {'value' : 'QA'},         # fixed - Found in Environment
            'customfield_15032' : {'value' : 'Production'}, # fixed - Found in Phase
            'labels' : ['L3'],                              # fixed
            'priority' : {'name' : 'Minor'},                # fixed
            #'Activity' : 'QA',
            'customfield_10170' : 'SPi',                    # fixed SPi
            'summary': 'Testing PM ',  # variable
            'description': 'Testing for accuracy',          # variable
            'customfield_11330': 'QUA-34',                  # variable - Epic link
            'customfield_10171' : 'ISBN test',              # variable - ISBN
            'customfield_10172' : 'DemandID',               # variable - DemandId
            'customfield_10202' : 'Author',                 # variable - Author
            'customfield_10203' : 'chapter',                # variable - Chapter
            }

"""
import sys
import csv
import time
import colorama
from jira import JIRA


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


def read_jira_info(filename, header):
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


def create_jira_issue(jira, jira_list, jira_update_list):
    """ update JIRA with all the new issues """
    final_map = {"Project": 0, "Issue Type": 1, "Summary": 2, "Description": 3, "Epic Link": 4, "Severity": 5,
                 "Reporter": 6, "Found in Environment": 7, "Found in Phase": 8, "Labels": 9, "Priority": 10,
                 "Activity": 11, "Custom1": 12, "Custom2": 13, "Custom3": 14, "Custom4": 15, "Custom5": 16,
                 "Edition": 17, "Primary": 18, "Title": 19}
    del jira_list[0]
    # change next line to 1 for production
    # TODO
    production = 0
    entry_index = 1
    for line in jira_list:
        issue_entry = {}
        issue_entry.update({'project': 'QUA'})
        issue_entry.update({'issuetype': {'name': 'Bug'}})
        issue_entry.update({'customfield_10080': {'value': 'Medium'}})
        issue_entry.update({'customfield_10061': {'value': 'QA'}})
        issue_entry.update({'customfield_15032': {'value': 'Production'}})
        issue_entry.update({'labels': ['L3']})
        issue_entry.update({'priority': {'name': 'Minor'}})
        # TODO uncomment this for production
        # issue_entry.update({'Activity' : 'QA',})
        issue_entry.update({'customfield_10170': 'SPi'})
        # TODO uncomment this for production
        # issue_entry.update({'reporter' : line[get_index_by_column_name("Reporter", final_map)]})
        issue_entry.update({'summary': line[get_index_by_column_name("Summary", final_map)]})
        issue_entry.update({'description': line[get_index_by_column_name("Description", final_map)]})
        issue_entry.update({'customfield_11330': line[get_index_by_column_name("Epic Link", final_map)]})
        issue_entry.update({'customfield_10171': line[get_index_by_column_name("Custom2", final_map)]})
        issue_entry.update({'customfield_10172': line[get_index_by_column_name("Custom3", final_map)]})
        issue_entry.update({'customfield_10202': line[get_index_by_column_name("Custom4", final_map)]})
        issue_entry.update({'customfield_10203': line[get_index_by_column_name("Custom5", final_map)]})
        if production:
            issue = jira.create_issue(fields=issue_entry)
            print(Bcolors.OKGREEN + f"{entry_index} - Created JIRA issue = '{issue}' Reporter = '{line[get_index_by_column_name('Reporter', final_map)]}' Primary = '{line[get_index_by_column_name('Primary', final_map)]}' Summary = '{line[get_index_by_column_name('Summary', final_map)]}'" + Bcolors.ENDC)
            jira.transition_issue(issue, 'Fixed', fields={'customfield_13831': {'value': 'Operator Error'}})
            print(Bcolors.OKGREEN + f"{entry_index} - Changed JIRA issue {issue} to Fixed " + Bcolors.ENDC)
            jira.transition_issue(issue, 'Close')
            print(Bcolors.OKGREEN + f"{entry_index} - Changed JIRA issue {issue} to Closed " + Bcolors.ENDC)
        else:
            print(Bcolors.OKGREEN + f"{entry_index} - Emulating - Created JIRA issue = 'QUA-{entry_index}',  Reporter = '{line[get_index_by_column_name('Reporter', final_map)]}'" + Bcolors.ENDC)
            print(Bcolors.OKGREEN + f"{entry_index} - Emulating - Changed JIRA issue = 'QUA-{entry_index}' to Fixed " + Bcolors.ENDC)
            print(Bcolors.OKGREEN + f"{entry_index} - Emulating - Changed JIRA issue = 'QUA-{entry_index}' to Closed " + Bcolors.ENDC)
        jira_update_list.append(f"{entry_index} - Created JIRA issue = 'QUA-{entry_index}'")
        time.sleep(1)
        entry_index += 1
        if not production:
            if entry_index == 5:
                break


def get_all_open_jira_issues(jira, username):
    """ see if there's any Open issues """
    total = 0
    initial = 0
    size = 50
    open_list = []
    while True:
        start = initial*size
        issues = jira.search_issues('project=CSC_Pearson_Quality and status != Closed and type=Bug', start, size)
        for issue in issues:
            if (str(issue.fields.status) != 'Open') and (issue.fields.creator.name == username):
                print(Bcolors.OKGREEN + f"JIRA issue key = '{issue}' Status = {issue.fields.status} Creator = {issue.fields.creator.name}" + Bcolors.ENDC)
                open_list.append(issue)
        initial += 1
        if not issues:
            break
        total += len(issues)


def get_all_jira_epics(programinput_path, user_credentials):
    """ update all Epic files with latest information from JIRA """
    print(30 * "-" + Bcolors.OKBLUE + "Retrieving Epics" + Bcolors.ENDC + 30 * "-")
    options = {'server': 'https://agile-jira.pearson.com'}
    try:
        jira = JIRA(options, basic_auth=(user_credentials[1], user_credentials[2]))
    except Exception:
        print(Bcolors.FAIL + "JIRA credentials not valid. Please delete file 'credentials.txt' and restart program" + Bcolors.ENDC)
        sys.exit(1)
    project_names = ['QUA', 'PM', 'PRODN']
    # create a list of of issues from JIRA
    for project_name in project_names:
        initial = 0
        total = 0
        size = 50
        epics_issue_list = []
        while True:
            start = initial*size
            epics = jira.search_issues(f'project={project_name} and type=Epic', start, size)
            for epic in epics:
                epics_issue_list.append(epic)  # GET RID OF THIS IN PRODUCTION
            initial += 1
            print('.', end='', flush=True)
            if not epics:
                break
            total += len(epics)
        print(' ')
        filename = str(project_name).strip('[]')
        print(f'Total Epics retrieved = {total} from JIRA {filename} Project')
        # convert list of issues type to list of strings
        epics_list = []
        epics_list.append(['Issue Type', 'Issue key', 'Issue id', 'Summary', 'Custom field (CustomField3)', 'Reporter'])
        for epic in epics_issue_list:
            if epic.key is None:
                epickey = ''
            else:
                epickey = epic.key
            # summary field
            if epic.fields.customfield_11331 is None:
                custom11331 = ''
            else:
                custom11331 = epic.fields.customfield_11331
            # Custom field3
            if epic.fields.customfield_10172 is None:
                custom10172 = ''
            else:
                custom10172 = epic.fields.customfield_10172
            epics_list.append([' ', epickey, ' ', custom11331, custom10172, ' '])
        # output list to a file
        filename = programinput_path + 'Epics' + str(project_name).strip('[]') + '.txt'
        with open(filename, 'w') as file_hndl:
            try:
                file_hndl.writelines('\t'.join(i) + '\n' for i in epics_list)
            except IOError:
                print(Bcolors.FAIL + "Cannot open file:" + filename + Bcolors.ENDC)
                sys.exit(1)


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


def access_jira(jira_input_path, user_credentials):
    """ main access point for routines """
    colorama.init()
    filename = jira_input_path + 'JIRAImportData.txt'
    jira_list = read_jira_info(filename, 1)
    # create JIRA class object
    jira_update_list = []
    options = {'server': 'https://agile-jira.pearson.com'}
    try:
        jira = JIRA(options, basic_auth=(user_credentials[1], user_credentials[2]))
    except Exception:
        print(Bcolors.FAIL + "JIRA credentials not valid. Please delete file 'credentials.txt' and restart program" + Bcolors.ENDC)
        sys.exit(1)
    print(30 * "-" + Bcolors.OKBLUE + "Creating JIRA Issues" + Bcolors.ENDC + 30 * "-")
    create_jira_issue(jira, jira_list, jira_update_list)
    print(30 * "-" + Bcolors.OKBLUE + "Retrieving Open JIRA Issues" + Bcolors.ENDC + 30 * "-")
    get_all_open_jira_issues(jira, user_credentials[1])
    filename = jira_input_path + 'JIRAUpdates.txt'
    write_updates_to_file(filename, jira_update_list)




