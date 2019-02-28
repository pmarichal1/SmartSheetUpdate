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
from __future__ import absolute_import
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


def get_smartsheet_ids(program_input_path):
    '''go get the Smartsheet IDs we want to download'''
    try:
        ss_hndl = open(program_input_path + "smartsheetGetIDs.txt", 'r')
        line_list = ss_hndl.readlines()
        #strip off newlines
        ids_list = []
        for line in line_list:
            ids_list.append(line.rstrip())
        ss_hndl.close()
        return ids_list
    except IOError:
        print(Bcolors.FAIL + 'No such file: %s' % program_input_path +
              'smartsheetGetIDs.txt' + Bcolors.ENDC)
        sys.exit(1)


def get_access_token(program_input_path):
    '''go get the access token we need for CURL request'''
    try:
        token_hndl = open(program_input_path + "accessToken.txt", 'r')
        access_token = token_hndl.readline()
        access_token = access_token.rstrip()
        token_hndl.close()
        return access_token
    except IOError:
        print(Bcolors.FAIL + 'No such file: %s' % program_input_path +
              'accessToken.txt' + Bcolors.ENDC)
        sys.exit(1)


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
            tab_text = open(program_input_path + tab_filename, 'w')
            try:
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
                    txt_writer = csv.writer(tab_text, delimiter='\t') #writefile
                    txt_writer.writerow(my_row)   #write the lines to file`
            except IOError:
                print(Bcolors.FAIL + 'No such file: %s' % program_input_path + 
                      tab_text + Bcolors.ENDC)
                sys.exit(1)
            tab_text.close()


def getSmartsheetMain(program_input_path):
    ''' main function '''
    colorama.init()
    #program_input_path = sys.argv[1]
    print("Python Version is " + platform.python_version())
    print("System Version is " + platform.platform())
    #get files from Smartsheet as .xlsx and covert them to tab delimited .txt files
    access_token = get_access_token(program_input_path)
    smartsheet_ids = get_smartsheet_ids(program_input_path)
    ss_client = smartsheet.Smartsheet(access_token)
    for ids in smartsheet_ids:
        ss_client.Reports.get_report_as_excel(ids, program_input_path)
    convert_downloaded_files_to_tab_delimited(program_input_path)
    print("All Smartsheets Downloaded")
    colorama.deinit()

if __name__ == "__main__":
    getSmartsheetMain(program_input_path)
