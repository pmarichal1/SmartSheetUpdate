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
example curl request

os.system("curl https://api.smartsheet.com/2.0/reports/4320876536588164 -H
            ""Authorization: Bearer hr3uz76quvxy56oj2v9ofdcpxo"" -H ""Accept:
                text/xlsx"" -o test.xlsx > /dev/null 2>&1")
    curl -k -v https://api.smartsheet.com/2.0/reports/4320876536588164 -H ""Authorization:
        Bearer hr3uz76quvxy56oj2v9ofdcpxo"" -H ""Accept: text/xlsx"" -o 4320876536588164.xlsx
            -D 4320876536588164.txt > /dev/null 2>&1
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

class Bcolors(object):
    ''' define colors for printing '''
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
            ids_list.append(line.strip())
        ss_hndl.close()
        return ids_list
    except IOError:
        print (Bcolors.FAIL + 'No such file: %s' % program_input_path +
               'smartsheetGetIDs.txt' + Bcolors.ENDC)
        sys.exit(1)


def get_access_token(program_input_path):
    '''go get the access token we need for CURL request'''
    try:
        token_hndl = open(program_input_path + "accessToken.txt", 'r')
        access_token = token_hndl.readline()
        token_hndl.close()
        return access_token
    except IOError:
        print (Bcolors.FAIL + 'No such file: %s' % program_input_path +
               'accessToken.txt' + Bcolors.ENDC)
        sys.exit(1)


def process_curl_request(access_token, ids_list):
    '''build URL request'''
    str_url = 'curl -k https://api.smartsheet.com/2.0/reports/'
    #authentication and request for xlsx file
    str_url_args = ' -H "Authorization: Bearer ' + access_token + \
                   '" -H "application/vnd.ms-excel"'
    # get header information
    str_header = ' -D '
    #redirect output to be quiet

    for ids in ids_list:
        #process CURL request
        os.system("""""" + str_url + ids + str_url_args + ' -o ' + program_input_path + ids +
                  '.xlsx' + str_header + program_input_path + ids + '.txt' + """""")


def process_downloaded_files(smartsheet_ids):
    '''let's open each new downloaded file and search for filename'''
    for ids in smartsheet_ids:
        smartsheet_names = []
        try:
            with open(program_input_path + ids + '.txt', 'r') as file_list:
                # chunk the header info delimited by ""
                # quotes so we can get filename from ids that we downloaded
                # there only 1 set of quotes in the file so minimal iterations
                hdr_list = file_list.read().split('"')
                #the second list entry is always the filename from the CURL get header
                line = hdr_list[1]
                #get rid of the filename extension
                filename = line.rsplit(".", 1)[0]
                smartsheet_names.append(program_input_path + filename)
                file_list.close()
                try:
                    #remove the .xlsx file it exist from previous run.
                    #This is required for Windows to work correctly
                    if os.path.isfile(program_input_path + filename + '.xlsx'):
                        os.remove(program_input_path + filename + '.xlsx')
                    os.rename(program_input_path + ids + '.xlsx', program_input_path + filename + '.xlsx')
                except OSError:
                    print (Bcolors.FAIL + 'Cannot rename: %s' % program_input_path + ids + '.txt' + 'to' +
                           program_input_path + filename + '.txt' + Bcolors.ENDC)
                    sys.exit(1)
                    #can't remove temporary files in windows
                try:
                    os.remove(program_input_path + ids + '.txt')
                except OSError:
                    print (Bcolors.FAIL + 'Cannot remove: %s' % program_input_path + ids + '.txt' + 'to' +
                           program_input_path + filename + '.txt' + Bcolors.ENDC)
                    sys.exit(1)
                print  ('Smartsheet ' + Bcolors.HEADER + filename + '.xlsx ' +
                        Bcolors.ENDC + 'downloaded')
        except IOError:
            print (Bcolors.FAIL + 'No such file: %s' % ids + '.txt' + Bcolors.ENDC)
            sys.exit(1)


def convert_downloaded_files_to_tab_delimited(program_input_path):
    '''let's open each new downloaded file and search for filename'''
    for full_filename in os.listdir(program_input_path):  # use the directory name here
        filename, file_ext = os.path.splitext(full_filename)
        if file_ext == '.xlsx':
            wb = load_workbook(filename=program_input_path + full_filename)
            #first_sheet = wb.get_sheet_names()[0]
            first_sheet = wb.active
            #worksheet = wb.get_sheet_by_name(first_sheet)
            tab_filename = filename + '.txt'
            print ('Converting file ' + Bcolors.HEADER + filename + 
                '.xlsx' + Bcolors.ENDC + 
                ' to Tab delimited ' + Bcolors.OKGREEN + filename + '.txt' + Bcolors.ENDC)
            tab_text = open(program_input_path + tab_filename, 'w')
            try:
                for row in first_sheet.iter_rows():
                    my_row = []
                    for cell in row:
                        value = cell.internal_value
                        result = type (value) is float
                        if result:
                            value = int(value)
                        my_row.append(value)
                    txt_writer = csv.writer(tab_text, delimiter='\t') #writefile
                    txt_writer.writerow(my_row)   #write the lines to file`
            except IOError:
                            print(Bcolors.FAIL + 'No such file: %s' % program_input_path +
                                  tab_txt + Bcolors.ENDC)
                            sys.exit(1)
            tab_text.close()


def main():
    program_input_path = sys.argv[1]
    print ("Python Version is " + platform.python_version())
    print ("System Version is " + platform.platform())
    ''' get files from Smartsheet as .xlsx and covert them to tab delimited .txt files'''
    access_token = get_access_token(program_input_path)
    smartsheet_ids = get_smartsheet_ids(program_input_path)
    ss_client = smartsheet.Smartsheet(access_token)
    #process_curl_request(access_token, smartsheet_ids)
    for ids in smartsheet_ids:
        ss_client.Reports.get_report_as_excel(ids, program_input_path)
    #the call is needed if making CURL request
    #process_downloaded_files(smartsheet_ids)
    convert_downloaded_files_to_tab_delimited(program_input_path)
    print ("All Smartsheets Downloaded")


if __name__ == "__main__":
    main()
