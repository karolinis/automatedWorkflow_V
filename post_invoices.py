import time

import requests
import os
import pandas as pd
import tkinter as tk
import re

def main(journal,acc_year,HEADER_P):
    WT = pd.read_excel('working_file.xlsx')

    for i in WT["Entry"]:
        try:
            j = WT.loc[WT["Entry"] == i, "File"].iloc[0]
            files = {'data': open(f'{ CURDIR }/Upload/' + str(j),'rb')}
            purl = "https://restapi.e-conomic.com/journals-experimental/" + journal + "/vouchers/" + acc_year + "-" + str(i) + "/attachment/file"
            p = requests.post(purl, headers=HEADER_P, files=files)
            print(p.status_code)
            if p.status_code == 201:
                print("201 Created OK - entry " + str(i))
            else:
                print(f"Status {p.status_code}")
        except Exception as e:
            print(e)
            time.sleep(3)
            pass

class Inp:
    def bt_func(self):
        self.root.destroy()

    def __init__(self):
        self.root = tk.Tk()
        self.root.title('E-conomic File attacher')
        self.root.geometry('600x350')
        self.root['bg'] = 'black'

        self.journal_name = tk.Label(self.root, text='Journal:', bg='black', fg='white')
        self.journal_name.pack()
        self.journal_var = tk.StringVar()
        self.journal_path = tk.Entry(self.root, width=20, borderwidth=5, textvariable=self.journal_var)
        self.journal_path.pack(pady=15)

        self.accounting_year = tk.Label(self.root, text='Accounting year:', bg='black', fg='white')
        self.accounting_year.pack()
        self.accounting_var = tk.StringVar()
        self.accounting_path = tk.Entry(self.root, width=20, borderwidth=5, textvariable=self.accounting_var)
        self.accounting_path.pack(pady=15)

        self.agreement_token = tk.Label(self.root, text='Agreement token:', bg='black', fg='white')
        self.agreement_token.pack()
        self.agreement_var = tk.StringVar()
        self.agreement_path = tk.Entry(self.root, width=20, borderwidth=5, textvariable=self.agreement_var)
        self.agreement_path.pack(pady=15)

        self.inspection = tk.Label(self.root, text='Make sure your working_file is closed when submitting',
                                     bg='black', fg='white')
        self.inspection.pack()

        self.btn = tk.Button(self.root,text='Submit',bg='white', command=self.bt_func)
        self.btn.pack()

        self.root.mainloop()


if __name__ == '__main__':
    CURDIR = os.getcwd()
    CURDIR = re.sub('\\\\','/',CURDIR)

    e = Inp()
    journal = e.journal_var.get()
    acc_year = e.accounting_var.get()
    AGRTOKEN = e.agreement_var.get()

    header = {'X-AppSecretToken': 'lVODBpMxjJt6u60pmcxJ47oqyT0qE79PBaUfbNaXxqI1',
              'X-AgreementGrantToken': AGRTOKEN,
              "cache-control": "no-cache"
              }
    try:
        main(journal,acc_year,header)
    except Exception as e:
        print(e)
        time.sleep(3)
        main(journal,acc_year,header)