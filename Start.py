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


def get_smartsheets(INPUT_CODE_PATH, INPUT_PROGRAMINPUT_PATH):
    getSmartsheetMain(INPUT_PROGRAMINPUT_PATH)


def create_one_jira_import_file(MAC_OR_WIN, INPUT_CODE_PATH):
    ''' create JIRA single import file '''
    if MAC_OR_WIN == '0':
        os.system(INPUT_CODE_PATH + "macJIRAOne")
    else:
        os.system(INPUT_CODE_PATH + "winJIRAOne.bat")


def create_multi_jira_import_files(MAC_OR_WIN, INPUT_CODE_PATH):
    ''' create JIRA multiple import files '''
    if MAC_OR_WIN == '0':
        os.system(INPUT_CODE_PATH + "macJIRAmultiple")
    else:
        os.system(INPUT_CODE_PATH + "winJIRAmultiple.bat")        

def update_smartsheets(INPUT_CODE_PATH, INPUT_PROGRAMINPUT_PATH, INPUT_JIRAINPUT_PATH):
    ''' update smartsheet in server '''
    while True:
        print("Are you sure you want to update Smartsheets?")
        try:
            # Note: Python 2.x users should use raw_input, the equivalent of 3.x's input
            if sys.version_info.major < 3:
                choice = str(raw_input("yes or no: "))
            else:
                choice = str(input("yes or no: "))
        except ValueError:
            print("Not a valid selection")
        if choice == 'yes':
            print(Bcolors.OKGREEN + "Updating Smartsheets" + Bcolors.ENDC)
            updateSmartsheetMain(INPUT_PROGRAMINPUT_PATH, INPUT_JIRAINPUT_PATH)
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
    MAC_OR_WIN = sys.argv[1]
    INPUT_CODE_PATH = sys.argv[2]
    INPUT_PROGRAMINPUT_PATH = sys.argv[3]
    INPUT_JIRAINPUT_PATH = sys.argv[4]
    print("Python Version from is " + platform.python_version())
    print("System Version is " + platform.platform())

    while True:
        print_menu()  ## Displays menu
        try:
            # Note: Python 2.x users should use raw_input, the equivalent of 3.x's input
            if sys.version_info.major < 3:
                choice = int(raw_input("Enter your choice [1-4]: "))
            else:
                choice = int(input("Enter your choice [1-4]: "))
        except ValueError:
            print("Not a valid number selection")
            # better try again... Return to the start of the loop
            continue
        if choice < 1 or choice > 4:
            print(Bcolors.FAIL + "Selection not in range" + Bcolors.ENDC)
            continue

        if choice == 1:
            get_smartsheets(INPUT_CODE_PATH, INPUT_PROGRAMINPUT_PATH)
        elif choice == 2:
            create_one_jira_import_file(MAC_OR_WIN, INPUT_CODE_PATH)
        elif choice == 3:
            update_smartsheets(INPUT_CODE_PATH, INPUT_PROGRAMINPUT_PATH, INPUT_JIRAINPUT_PATH)
        elif choice == 4:
            colorama.deinit()
            break

if __name__ == "__main__":
    main()
