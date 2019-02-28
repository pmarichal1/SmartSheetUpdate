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
FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''

from __future__ import print_function
from __future__ import unicode_literals
from __future__ import absolute_import
import platform
import sys
import os
import colorama
from getSmartsheets import getSmartsheetMain
from updateSmartsheets import updateSmartsheetMain

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


def get_smartsheets(input_programinput_path):
    '''call getSmaartheet module'''
    getSmartsheetMain(input_programinput_path)


def create_one_jira_import_file(mac_or_win, input_code_path):
    ''' create JIRA single import file '''
    if mac_or_win == '0':
        temp_path = '"' + input_code_path + "macJIRAOne" + '"'
        os.system(temp_path)
    else:
        temp_path = input_code_path + "winJIRAOne.bat"
        os.system('"' + temp_path + '"')


def create_multi_jira_import_files(mac_or_win, input_code_path):
    ''' create JIRA multiple import files'''
    if mac_or_win == '0':
        temp_path = '"' + input_code_path + "macJIRAmultiple" + '"'
        os.system(temp_path) 
    else:
        temp_path = input_code_path + "winJIRAmultiple.bat"
        os.system('"' + temp_path + '"')



def update_smartsheets(input_programinput_path, input_jirainput_path):
    ''' update smartsheet in server '''
    while True:
        print("Are you sure you want to update Smartsheets?")
        try:
            choice = str(input("yes or no: "))
        except ValueError:
            print("Not a valid selection")
        if choice == 'yes':
            print(Bcolors.OKGREEN + "Updating Smartsheets" + Bcolors.ENDC)
            updateSmartsheetMain(input_programinput_path, input_jirainput_path)
            return
        elif choice == 'no':
            print(Bcolors.OKGREEN  + "Smartsheets not updated" + Bcolors.ENDC)
            return
        else:
            continue


def print_menu():
    ''' print initial menu for user'''
    print(30 * "-" + Bcolors.OKBLUE + "MENU" + Bcolors.ENDC + 30 * "-")
    print(Bcolors.OKGREEN + "1   - Get Updated JIRA Smartsheets")
    print("2   - Create JIRA file to be Imported")
    print("3   - Update Smartsheets from JIRA Import File")
    print("4   - Quit" + Bcolors.ENDC)
    print(67 * "-")


def main():
    ''' main loop to get us started'''
    colorama.init()
    # 0 to invoke mac script or 1 for windows
    mac_or_win = sys.argv[1]
    input_code_path = sys.argv[2]
    input_programinput_path = sys.argv[3]
    input_jirainput_path = sys.argv[4]
    print("Python Version from is " + platform.python_version())
    print("System Version is " + platform.platform())
    print("Sotfware Version is V3.3.0")

    while True:
        print_menu()  ## Displays menu
        try:
            choice = int(input("Enter your choice [1-4]: "))
        except ValueError:
            print("Not a valid number selection")
            # better try again... Return to the start of the loop
            continue
        if choice < 1 or choice > 4:
            print(Bcolors.FAIL + "Selection not in range" + Bcolors.ENDC)
            continue

        if choice == 1:
            get_smartsheets(input_programinput_path)
        elif choice == 2:
            create_one_jira_import_file(mac_or_win, input_code_path)
        elif choice == 3:
            update_smartsheets(input_programinput_path, input_jirainput_path)
        elif choice == 4:
            colorama.deinit()
            break

if __name__ == "__main__":
    main()
