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
    if file:
        typewrite('Connecting to sheets', '...')
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('secret_access.json', scope)
        client = gspread.authorize(creds)
        typewrite('Openning Hilmar 2', '...')
        hilmar2 = get_as_dataframe(client.open('CIT 2020 LogSheet').worksheet(title="Hilmar 2"))
        hilmar2['Delivered'] = hilmar2['Delivered'].fillna(0)
        hilmar2 = hilmar2[(hilmar2['Delivered'] != 0)]
        hilmar2_no_c = hilmar2[(hilmar2['Status'] != 'C')]
        typewrite('Openning report', '...')
        report = pd.read_excel(file)
        typewrite('Converting values to integers', '...')
        temp = list()
        for item in report['Load ID']:
            temp.append(item)

        report['Load ID'] = temp
        found_df = report[~report['Load ID'].isin(hilmar2['Delivered'])]
        found_df_no_c = report[report['Load ID'].isin(hilmar2_no_c['Delivered'])]
        report_filter = report[report['Load ID'].isin(list(found_df['Load ID']) +
                                                      list(found_df_no_c['Load ID']))]
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
    else:
        typewrite("No file selected", "...")


canvas = tk.Canvas(root, height=10, width=200)
canvas.pack()

generateReport = tk.Button(root, text="Generate report", command=genReport)
generateReport.pack()

root.mainloop()
