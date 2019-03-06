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
#from __future__ import print_function
#from __future__ import absolute_import
import sys
import os
import csv
import platform
import smartsheet
from smartsheet import reports
from openpyxl import load_workbook
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


def get_file_info(filename):
    '''go get the file information'''
    try:
        ss_hndl = open(filename, 'r')
    except FileNotFoundError as e:
        print(Bcolors.FAIL + f"File not found {filename}" + Bcolors.ENDC)
        sys.exit(1)
    except Exception as e:
        errno, strerror = e.args
        print(Bcolors.FAIL + f"I/O error({errno}): {strerror} {filename}" + Bcolors.ENDC)
        sys.exit(1)
    else:
        line_list = ss_hndl.readlines()
        #strip off newlines
        info_list = []
        for line in line_list:
            info_list.append(line.rstrip())
        ss_hndl.close()
        return info_list


def convert_downloaded_files_to_tab_delimited(program_input_path):
    '''let's open each new downloaded file and search for filename'''
    for full_filename in os.listdir(program_input_path):  # use the directory name here
        filename, file_ext = os.path.splitext(full_filename)
        if file_ext == '.xlsx':
            workbook = load_workbook(filename=program_input_path + full_filename)
            first_sheet = workbook.active
            tab_filename = filename + '.txt'
            print('Converting file ' + Bcolors.HEADER + filename +
                  '.xlsx' + Bcolors.ENDC + ' to Tab delimited ' +
                  Bcolors.OKGREEN + filename + '.txt' + Bcolors.ENDC)
            try:
                tab_text = open(program_input_path + tab_filename, 'w')
            except Exception as e:
                errno, strerror = e.args
                print(Bcolors.FAIL + f"I/O error({errno}): {strerror} {program_input_path}{tab_text}" + Bcolors.ENDC)
                sys.exit(1)
            else:
                txt_writer = csv.writer(tab_text, delimiter='\t', lineterminator='\n')  # writefile
                for row in first_sheet.iter_rows():
                    my_row = []
                    for cell in row:
                        value = cell.internal_value
                        result = isinstance(value, str)
                        if result:
                            if '\n' in value:
                                value = value.replace('\n', ' ')
                        result = isinstance(value, float)
                        if result:
                            value = int(value)
                        my_row.append(value)
                    txt_writer.writerow(my_row)   #write the lines to file`
                tab_text.close()


def getSmartsheetMain(program_input_path):
    ''' main function '''
    colorama.init()
    #program_input_path = sys.argv[1]
    print("Python Version is " + platform.python_version())
    print("System Version is " + platform.platform())
    #get files from Smartsheet as .xlsx and covert them to tab delimited .txt files
    filename = program_input_path + "accessToken.txt"
    #access_token = get_access_token(filename)
    access_token = get_file_info(filename)
    access_token = access_token[0]
    filename = program_input_path + "smartsheetGetIDs.txt"
    smartsheet_ids = get_file_info(filename)
    ss_client = smartsheet.Smartsheet(access_token)
    for ids in smartsheet_ids:
        try:
            ss_client.Reports.get_report_as_excel(ids, program_input_path)
        except Exception:
            print(f"You Token is not valid in accessToken.txt")
            sys.exit(1)
    convert_downloaded_files_to_tab_delimited(program_input_path)
    print("All Smartsheets Downloaded")
    colorama.deinit()

if __name__ == "__main__":
    getSmartsheetMain(program_input_path)
