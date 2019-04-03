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
FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import base64
import os
import platform
import sys
import time
from getpass import getpass

import colorama

from createJiraImport import create_jira_import_main
from getSmartsheets import getSmartsheetMain
from updateJira import access_jira, get_all_jira_epics
from updateSmartsheets import updateSmartsheetMain


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

def encode(key,  clear):
    """ encode information for security """
    enc = []
    for i in range(len(clear)):
        key_c = key[i % len(key)]
        enc_c = chr((ord(clear[i]) + ord(key_c)) % 256)
        enc.append(enc_c)
    return base64.urlsafe_b64encode("".join(enc).encode()).decode()


def decode(key, enc):
    """ decode information for security """

    dec = []
    enc = base64.urlsafe_b64decode(enc).decode()
    for i in range(len(enc)):
        key_c = key[i % len(key)]
        dec_c = chr((256 + ord(enc[i]) - ord(key_c)) % 256)
        dec.append(dec_c)
    return "".join(dec)


def encode_user_credentials(filename, crypt_key, token, username, password):
    credentials = []
    crypt_token = encode(crypt_key, token)
    crypt_username = encode(crypt_key, username)
    crypt_password = encode(crypt_key, password)
    credentials.append(crypt_token)
    credentials.append(crypt_username)
    credentials.append(crypt_password)
    try:
        f = open(filename, 'w')
    except FileNotFoundError:
        print(Bcolors.FAIL + "Cannot open file:" + filename + Bcolors.ENDC)
        sys.exit(1)
    except Exception as e:
        errno, strerror = e.args
        print(Bcolors.FAIL + f"I/O error {errno}: {strerror} {filename}" + Bcolors.ENDC)
        sys.exit(1)
    else:
        for line in credentials:
            f.write(line + '\n')
        f.close()


def decode_user_credentials(filename, crypt_key):
    clear_list = []
    try:
        f = open(filename, 'r')
    except Exception as e:
        errno, strerror = e.args
        print(Bcolors.FAIL + f"I/O error({errno}): {strerror} {filename}" + Bcolors.ENDC)
        sys.exit(1)
    else:
        for line in f.readlines():
            try:
                clear_list.append(decode(crypt_key, line))
            except Exception as e:
                print(Bcolors.FAIL + "credentials.txt is possibly corrupted. Please delete file 'credentials.txt' and restart program" + Bcolors.ENDC)
        f.close()
        return clear_list

def update_smartsheets(input_programinput_path, input_jirainput_path, user_credentials):
    """ update smartsheet in server """
    while True:
        print("Are you sure you want to update Smartsheets?")
        try:
            choice = str(input("yes or no: "))
        except ValueError:
            print("Not a valid selection")
        except Exception as e:
            print(f"{e}")
        else:
            if choice == 'yes':
                print(Bcolors.OKGREEN + "Updating Smartsheets" + Bcolors.ENDC)
                updateSmartsheetMain(input_programinput_path, input_jirainput_path, user_credentials)
                return
            elif choice == 'no':
                print(Bcolors.OKGREEN + "Smartsheets not updated" + Bcolors.ENDC)
                return
            else:
                continue

def access_jira_local(jirainput_path, user_credentials, production):
    """ update JIRA in server """
    while True:
        print("Are you sure you want to update JIRA?")
        try:
            choice = str(input("yes or no: "))
        except ValueError:
            print("Not a valid selection")
        except Exception as e:
            print(f"{e}")
        else:
            if choice == 'yes':
                print(Bcolors.OKGREEN + "Updating JIRA" + Bcolors.ENDC)
                access_jira(jirainput_path, user_credentials, production)
                return
            elif choice == 'no':
                print(Bcolors.OKGREEN + "JIRA not updated" + Bcolors.ENDC)
                return
            else:
                continue

def print_menu():
    """ print initial menu for user"""
    print(30 * "-" + Bcolors.OKBLUE + "MENU" + Bcolors.ENDC + 30 * "-")
    print(Bcolors.OKGREEN + "1   - Generate JIRA Import File")
    print("2   - Create Issues and Update Workflow in JIRA")
    print("3   - Update Smartsheets with Uploaded JIRA Info")
    print("4   - (Optional) Update Local Epic Lookup Files")
    print("5   - Quit" + Bcolors.ENDC)
    print(67 * "-")

def main():
    """ main loop to get us started"""
    colorama.init()
    programinput_path = sys.argv[1]
    jirainput_path = sys.argv[2]
    production = sys.argv[3]
    print("Python Version from is " + platform.python_version())
    print("System Version is " + platform.platform())
    print("Software Version is V5.6.1")
    localtime = time.asctime(time.localtime(time.time()))
    print("Local current time :", localtime)
    crypt_key = 'secret SPi message'
    filename = programinput_path + "credentials.txt"
    if os.path.isfile(filename):
        user_credentials = decode_user_credentials(filename, crypt_key)
    else:
        token = str(input("Please enter your Smartsheet Token: "))
        jira_username = str(input("Please enter your JIRA username (V42): "))
        jira_password = getpass("Please enter your JIRA password: ")
        encode_user_credentials(filename, crypt_key, token, jira_username, jira_password)
        user_credentials = decode_user_credentials(filename, crypt_key)

    while True:
        print_menu()  ## Displays menu
        try:
            choice = int(input("Enter your choice [1-5]: "))
        except ValueError:
            print("Not a valid number selection")
            # better try again... Return to the start of the loop
            continue
        if choice < 1 or choice > 5:
            print(Bcolors.FAIL + "Selection not in range" + Bcolors.ENDC)
            continue

        if choice == 1:
            getSmartsheetMain(programinput_path, user_credentials)
            create_jira_import_main(programinput_path, jirainput_path)
        elif choice == 2:
            access_jira_local(jirainput_path, user_credentials, production)
        elif choice == 3:
            update_smartsheets(programinput_path, jirainput_path, user_credentials)
        elif choice == 4:
            get_all_jira_epics(programinput_path, user_credentials)
        elif choice == 5:
            colorama.deinit()
            break


if __name__ == "__main__":
    main()
