import imaplib
import getpass
import email
import datetime
import re
import smtplib

#opening the excel sheet

from openpyxl import *
wb = load_workbook(filename='mbrs.xlsx', read_only=False)
ws = wb.active  # ws is now an IterableWorksheet

#logging in

user = raw_input("GMail :- ")
pwd = getpass.getpass()
 
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(user,pwd)
mail.list()
mail.select("inbox")



for col in ws.iter_cols(min_row=1,max_col=1, max_row = 3):
    for cell in col:
        mbr = str(cell.value)
        print("%s"%(mbr))
        
        result, data = mail.search(None, '(FROM mbr)')
        #result, data = mail.search(None, '(FROM "jeskris0@gmail.com")')
        ids = data[0] # data is a list.
        id_list = ids.split() # ids is a space separated string
        latest_email_id = id_list[-1]#get the latest 
        result, data = mail.fetch(latest_email_id, "RFC822") # fetch the email body (RFC822) for the given ID
        raw_email = data[0][1] 
        email_message = email.message_from_string(raw_email)
        print email.utils.parseaddr(email_message['From'])
        print email_message.items()
