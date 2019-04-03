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
from tkinter import Tk, Label, Button, messagebox, ttk, HORIZONTAL, Frame, PhotoImage, Entry, LabelFrame,Text, Scrollbar, END
import threading
import base64
import platform
import os
import sys
import time
import colorama


from createJiraImport import create_jira_import_main
from getSmartsheets import getSmartsheetMain
from updateJira import access_jira, get_all_jira_epics
from updateSmartsheets import updateSmartsheetMain


#class Bcolors:
#    """color palette for print statements"""
#    HEADER = '\033[96m'
#    OKBLUE = '\033[94m'
#    OKGREEN = '\033[92m'
#    ENDC = '\033[0m'
#    BOLD = '\033[1m'
#    UNDERLINE = '\033[4m'
#    WARNING = '\033[93m'
#    FAIL = '\033[91m'

VERSION = "JIRA Automation V5.9.0"

class Bcolors:
    """color palette for print statements"""
    HEADER = ''
    OKBLUE = ''
    OKGREEN = ''
    ENDC = ''
    BOLD = ''
    UNDERLINE = ''
    WARNING = ''
    FAIL = '\n\n ***************************************************   Error  ***************************************************\n'


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
    """ take clear text and encode """
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
        print(Bcolors.FAIL + f"I/O error {errno} : {strerror} {filename}" + Bcolors.ENDC)
        sys.exit(1)
    else:
        for line in credentials:
            f.write(line + '\n')
        f.close()


def decode_user_credentials(filename, crypt_key):
    """ take encrypted text and return clear text """

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
                print(e)
        f.close()
        return clear_list


class Std_redirector(object):
    def __init__(self,widget):
        self.widget = widget

    def write(self,string):
        self.widget.insert(END,string)
        self.widget.see(END)

    def flush(self):
        pass

class Authenticate(Tk):
    """docstring for Values"""

    def __init__(self, parent):
        Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.grid()
        self.frame1 = LabelFrame(self, text="Enter Credentials and Token ", labelanchor='n')
        self.frame1.grid(row=0, columnspan=7, sticky='W',padx=5, pady=5, ipadx=5, ipady=5)
        self.Val1Lbl = Label(self.frame1,text="JIRA Username")
        self.Val1Lbl.grid(row=0, column=0, sticky='E', padx=5, pady=2)
        self.Val1Txt = Entry(self.frame1)
        self.Val1Txt.grid(row=0, column=1, columnspan=3, pady=2, sticky='WE')
        self.Val2Lbl = Label(self.frame1,text="JIRA Password")
        self.Val2Lbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)
        self.Val2Txt = Entry(self.frame1, show="*")
        self.Val2Txt.grid(row=1, column=1, columnspan=3, pady=2, sticky='WE')
        self.Val3Lbl = Label(self.frame1, text="Smartsheet Token")
        self.Val3Lbl.grid(row=2, column=0, sticky='E', padx=5, pady=2)
        self.Val3Txt = Entry(self.frame1, show="*")
        self.Val3Txt.grid(row=2, column=1, columnspan=3, pady=2, sticky='WE')

        self.val1 = None
        self.val2 = None
        self.val3 = None

        SubmitBtn = Button(self.frame1, text="Submit",command=self.submit)
        SubmitBtn.grid(row=4, column=3, sticky='W', padx=5, pady=2)

    def submit(self):
        self.val1=self.Val1Txt.get()
        if self.val1=="":
            Win2=Tk()
            Win2.withdraw()

        self.val2=self.Val2Txt.get()
        if self.val2=="":
            Win2=Tk()
            Win2.withdraw()

        self.val3 = self.Val3Txt.get()
        if self.val3 == "":
            Win2 = Tk()
            Win2.withdraw()
        self.quit()


class MySPiGUI:
    """ tkinter class for GUI """
    def __init__(self, master, user_credentials, param_list):
        self.master = master
        self.user_credentials = user_credentials
        self.programinput_path = param_list[0]
        self.jirainput_path = param_list[1]
        self.production = param_list[2]
        self.click_png_path = param_list[0] + 'click.png'
        self.spi_png_path = param_list[0] + 'spi.png'
        self.master.tk.call('wm', 'iconphoto', self.master._w, PhotoImage(file=self.click_png_path))
        master.title(VERSION)

        self.image = PhotoImage(file=self.spi_png_path)
        self.label = Label(image=self.image)
        self.label.pack()

        self.frame1 = Frame(master, highlightbackground="black", highlightcolor="black", highlightthickness=4, bd= 2)
        self.frame1.pack()
        # bg='#e6f2ff',
        self.label = Label(self.frame1, width=30, font=("Calibri", 15), fg='blue', text="Click below to make a selection")
        self.label.pack()

        self.gen_button = Button(self.frame1, highlightthickness=5, width=40, font=("Calibri", 16), text="Step 1 - Download and Generate JIRA Import File       ", fg='black', command=self.func1)
        self.gen_button.pack()
        self.gen_button.bind("<Enter>", lambda event: self.gen_button.configure(fg="orange"))
        self.gen_button.bind("<Leave>", lambda event: self.gen_button.configure(fg="black"))

        self.create_button = Button(self.frame1, highlightthickness=5, width=40, font=("Calibri", 16), text="Step 2 - Create Issues and Update Workflows in JIRA", fg='black', command=self.func2)
        self.create_button.pack()
        self.create_button.bind("<Enter>", lambda event: self.create_button.configure(fg="orange"))
        self.create_button.bind("<Leave>", lambda event: self.create_button.configure(fg="black"))

        self.update_button = Button(self.frame1, highlightthickness=5, width=40, font=("Calibri", 16), text="Step 3 - Update Smartsheets with Uploaded JIRA Info", fg='black', command=self.func3)
        self.update_button.pack()
        self.update_button.bind("<Enter>", lambda event: self.update_button.configure(fg="orange"))
        self.update_button.bind("<Leave>", lambda event: self.update_button.configure(fg="black"))

        self.optional_button = Button(self.frame1, highlightthickness=5, font=("Calibri", 16), text="(Optional) Update Local Epic Lookup Files", fg='gray', command=self.func4)
        self.optional_button.pack()
        self.optional_button.bind("<Enter>", lambda event: self.optional_button.configure(fg="orange"))
        self.optional_button.bind("<Leave>", lambda event: self.optional_button.configure(fg="black"))

        self.progress = ttk.Progressbar(self.frame1, orient=HORIZONTAL, length=450, mode='determinate')
        self.progress.pack()

        self.close_button = Button(self.frame1, highlightthickness=5, width=10, text="Quit", fg='black', command=master.quit)
        self.close_button.pack()
        self.close_button.bind("<Enter>", lambda event: self.close_button.configure(fg="red"))
        self.close_button.bind("<Leave>", lambda event: self.close_button.configure(fg="black"))

        # create new frame that will contain output text frame with scrollbar
        self.frame2 = Frame(master,  highlightbackground="black", highlightcolor="black", highlightthickness=4, bd= 2)
        self.frame2.pack(expand=1)

        # bg='#e6f2ff',
        self.label = Label(self.frame2, width=90, font=("Calibri", 15), fg='blue', text="Text Output")
        self.label.pack(expand=1)

        # create a Scrollbar and associate it with txt
        self.scrollb = Scrollbar(self.frame2)
        self.scrollb.pack(side='right', fill='y')

        # create a Text widget
        self.txt = Text(self.frame2, width=115, font=("Calibri", 12), borderwidth=3, wrap='word', undo=True, yscrollcommand=self.scrollb.set)
        self.txt.pack(side='left', fill='both')
        self.scrollb.config(command=self.txt.yview)

    def func1(self):
        """ handle step1 functions """
        def step1_thread():
            retok = getSmartsheetMain(self.programinput_path, self.user_credentials)
            if not retok:
                self.progress.stop()
                self.close_button['state'] = 'normal'
                return(0)
            retok = create_jira_import_main(self.programinput_path, self.jirainput_path)
            if not retok:
                self.progress.stop()
                self.close_button['state'] = 'normal'
                return(0)
            messagebox.showinfo("Information", "Step 1 Completed Successfully")
            self.progress.stop()
            self.enable_buttons()

        threading.Thread(target=step1_thread).start()
        self.disable_buttons()
        self.progress.start(100)

    def func2(self):
        """ handle step2 functions """
        def step2_thread():
            retok = access_jira(self.jirainput_path, self.user_credentials, self.production)
            if not retok:
                self.progress.stop()
                self.close_button['state'] = 'normal'
                return(0)
            messagebox.showinfo("", "Step 2 Completed Successfully")
            self.progress.stop()
            self.enable_buttons()

        result = messagebox.askyesno("JIRA Update", "Are you sure you want to update JIRA?")
        if result is True:
            threading.Thread(target=step2_thread).start()
            self.disable_buttons()
            self.progress.start(100)

    def func3(self):
        """ handle step3 functions """
        def step3_thread():
            retok = updateSmartsheetMain(self.programinput_path, self.jirainput_path, self.user_credentials)
            if not retok:
                self.progress.stop()
                self.close_button['state'] = 'normal'
                return(0)
            messagebox.showinfo("", "Step 3 Completed Successfully")
            self.progress.stop()
            self.enable_buttons()

        result = messagebox.askyesno("JIRA Update", "Are you sure you want to update Smartsheets?")
        if result is True:
            threading.Thread(target=step3_thread).start()
            self.disable_buttons()
            self.progress.start(100)

    def func4(self):
        """ handle step4 functions """
        def step4_thread():
            retok = get_all_jira_epics(self.programinput_path, self.user_credentials)
            if not retok:
                self.progress.stop()
                self.close_button['state'] = 'normal'
                return(0)
            messagebox.showinfo("", "Optional Step Completed Successfully")
            self.progress.stop()
            self.enable_buttons()

        threading.Thread(target=step4_thread).start()
        self.disable_buttons()
        self.progress.start(100)

    def disable_buttons(self):
        """ make sure to disable button while processing """
        self.gen_button['state'] = 'disabled'
        self.create_button['state'] = 'disabled'
        self.update_button['state'] = 'disabled'
        self.optional_button['state'] = 'disabled'
        self.close_button['state'] = 'disabled'

    def enable_buttons(self):
        """ make sure to enabled button after processing """
        self.gen_button['state'] = 'normal'
        self.create_button['state'] = 'normal'
        self.update_button['state'] = 'normal'
        self.optional_button['state'] = 'normal'
        self.close_button['state'] = 'normal'


def main():
    """ main function called by script """
    param_list = []
    colorama.init()
    programinput_path = sys.argv[1]
    jirainput_path = sys.argv[2]
    production = sys.argv[3]
    param_list.append(programinput_path)
    param_list.append(jirainput_path)
    param_list.append(production)
    print("Python Version from is " + platform.python_version())
    print("System Version is " + platform.platform())
    print(VERSION)
    localtime = time.asctime(time.localtime(time.time()))
    print("Local current time :", localtime)
    crypt_key = 'secret SPi message'
    filename = programinput_path + "credentials.txt"
    if os.path.isfile(filename):
        user_credentials = decode_user_credentials(filename, crypt_key)
    else:
        app = Authenticate(None)
        app.title('JIRA Automation Authentication')
        app.mainloop()  # this will run until it closes
        jira_username = app.val1
        jira_password = app.val2
        token = app.val3
        app.destroy()
        encode_user_credentials(filename, crypt_key, token, jira_username, jira_password)
        user_credentials = decode_user_credentials(filename, crypt_key)

    root = Tk()
    root.geometry("1000x800")
    root.resizable(0, 0)  # Don't allow resizing in the x or y direction
    my_gui = MySPiGUI(root, user_credentials, param_list)
    # redirecting output from script to Tkinter Text window
    sys.stdout = Std_redirector(my_gui.txt)
    root.mainloop()
    # To stop redirecting stdout:
    sys.stdout = sys.__stdout__
    root.destroy()


if __name__ == "__main__":
    main()
