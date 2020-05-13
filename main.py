import os
import re
import sys
import shutil
import gspread
import openpyxl
import pandas as pd
import tkinter as tk
from time import sleep
from tkinter import filedialog
from gspread_dataframe import get_as_dataframe
from oauth2client.service_account import ServiceAccountCredentials


def typewrite(msg, dot):
    print(msg, end='')
    sleep(0.3)
    for c in dot:  # for each character in each line
        print(c, end='')  # print a single character, and keep the cursor there.
        sys.stdout.flush()  # flush the buffer
        sleep(0.3)  # wait a little to make the effect look good.
    print('')
    print('')


typewrite('Starting app', '...')


root = tk.Tk()


def addReport():
    file = filedialog.askopenfilename(initialdir="/", title="Select report",
                                      filetype=(("Excel Files", "*.xls"), ("All files", "*.*")))
    return file


def genReport():
    file = addReport()

    typewrite('Connecting to sheets', '...')
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('secret_access.json', scope)
    client = gspread.authorize(creds)

    typewrite('Openning Hilmar 2', '...')
    hilmar2 = get_as_dataframe(client.open('CIT 2020 LogSheet').worksheet(title="Hilmar 2"))
    hilmar2['Manifest'] = hilmar2['Manifest'].fillna(0)
    hilmar2 = hilmar2[(hilmar2['Manifest'] != 0)]
    hilmar2_no_c = hilmar2[(hilmar2['Status'] != 'C')]

    typewrite('Openning report', '...')
    report = pd.read_excel(file)

    typewrite('Converting values to integers', '...')
    temp = list()
    for item in report['Unnamed: 12']:
        splt = re.split(r'\s|\n|\t', item)
        for s in splt:
            if (len(s) >= 6) & (str.isdecimal(s)):
                temp.append(int(s))

    print(len(temp))

    print(len(report['Unnamed: 12']))

    report['Unnamed: 12'] = temp

    found_df = report[~report['Unnamed: 12'].isin(hilmar2['Manifest'])]
    found_df_no_c = report[report['Unnamed: 12'].isin(hilmar2_no_c['Manifest'])]
    report_filter = report[report['Unnamed: 12'].isin(list(found_df['Unnamed: 12']) +
                                                      list(found_df_no_c['Unnamed: 12']))]
    report_filter.to_excel('not_found.xlsx')
    dest = os.path.join(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'), 'not_found.xlsx')
    typewrite('Copynig file', '...')
    shutil.copyfile('not_found.xlsx', dest)
    print('File copied.')
    typewrite('Openning file', '...')
    os.startfile(dest)
    sleep(1)
    print('Done.')
    sleep(1)
    os.remove('not_found.xlsx')
    sys.exit()


canvas = tk.Canvas(root, height=10, width=200)
canvas.pack()

generateReport = tk.Button(root, text="Generate report", command=genReport)
generateReport.pack()

root.mainloop()
