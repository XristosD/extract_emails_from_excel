import re
import openpyxl
import os
import xlrd

path=r"C:\\Users\\xristos\\Desktop\\emex\\"

email = re.compile(r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$")

emails = []

for filename in os.listdir(path):
    if filename.endswith('.xls'):
        ws = xlrd.open_workbook(path+"pharmacy.xlsx").sheet_by_index(0)
        for x in range(ws.nrows):
            for f in ws.row(x):
                if email.match(f.value):
                    emails.append(f.value)

                
    elif filename.endswith('.xlsx'):
        ws = openpyxl.load_workbook(path+"pharmacy.xlsx").active
        for col in ws.rows:
            for cell in col:
                if type(cell.value) == str:
                    if email.match(cell.value):
                        emails.append(cell.value)
                    
                
emails_rem_dupl = []
emails.sort()    
    
for email in emails:
    if email not in emails_rem_dupl:
        emails_rem_dupl.append(email)
            
emails = emails_rem_dupl


        
