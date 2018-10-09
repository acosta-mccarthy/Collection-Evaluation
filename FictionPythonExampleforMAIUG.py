# -*- coding: utf-8 -*-
#!/usr/bin/python2.7
#
# Export shelf list to collection managers
# Email Excel Spreadsheet to manager and supervisor 
# Use XlsxWriter to create spreadsheet from SQL Query
# Use target variables to determine if collection size is above target
# Send one of two emails depending on rowcount

import psycopg2
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

excelfile =  'Fiction.xlsx'



#Set variables for email

emailhost = 'mail.greenwichlibrary.org'
emailport = '25'

try:
    conn = psycopg2.connect("dbname='iii' user='xxxxxxx' host='sierra-db' port='1032' password='xxxxxx' sslmode='require'")
except psycopg2.Error as e:
    print ("Unable to connect to database: " + str(e))
    
cursor = conn.cursor()
cursor.execute(open("Fiction.sql","r").read())
#Set variables to create targets
rows = cursor.fetchall()
target = cursor.rowcount 
target1 = target - 23279
emailsubject1 = '[Weeding] Fiction'
emailsubject2 = 'Fiction section on target!'
emailmessage1 = '''Hi!

You are above the target size of your collection. Please weed {0} items to meet target size of 23279 items. Attached you will find a shelf list containing circulation data.
If you have any questions, please contact Eric McCarthy in Resources Management. Currently the total size of the collection is {1}'''.format((target1),(target)) 
emailmessage2 = '''Hi!

Your collection is on or under target.  Currently the target is 23279 and there are {0} items in this collection.  Thanks!'''.format(target)
emailfrom = 'xxxxxxx@greenwichlibrary.org'
emailto = ['xxxxxxxx@greenwichlibrary.org','xxxxxxx@greenwichlibrary.org','xxxxxxx@greenwichlibrary.org']
conn.close()


workbook = xlsxwriter.Workbook(excelfile, {'remove_timezone': True})
worksheet = workbook.add_worksheet()

worksheet.set_landscape()
worksheet.hide_gridlines(0)
worksheet.repeat_rows(0)
worksheet.set_print_scale(58)
worksheet.set_margins(left=0.2, right=0.2)

eformat= workbook.add_format({'text_wrap': False, 'valign': 'bottom'})
eformat2= workbook.add_format({'text_wrap': False, 'valign': 'bottom', 'num_format': 'mm/dd/yy'})
eformatlabel= workbook.add_format({'text_wrap': False, 'valign': 'bottom', 'bold': True})

worksheet.set_column(0,0,14.43)
worksheet.set_column(1,1,14.43)
worksheet.set_column(2,2,62.86)
worksheet.set_column(3,3,21.57)
worksheet.set_column(4,4,8.71)
worksheet.set_column(5,4,13)
worksheet.set_column(6,6,18.57)
worksheet.set_column(7,7,9.29)
worksheet.set_column(8,8,9.29)
worksheet.set_column(9,9,4)
worksheet.set_column(10,10,4)
worksheet.set_column(11,11,11.71)
worksheet.set_column(12,12,15.86)
worksheet.set_column(13,13,8.43)

worksheet.set_header('&CFiction')
worksheet.set_footer('&CPage &P of &N&R&D')



worksheet.write(0,0,'Call Number', eformatlabel)
worksheet.write(0,1,'Item Call #', eformatlabel)
worksheet.write(0,2,'Title', eformatlabel)
worksheet.write(0,3,'Author', eformatlabel)
worksheet.write(0,4,'Pub. Year', eformatlabel)
worksheet.write(0,5,'Item Created', eformatlabel)
worksheet.write(0,6,'Last Checkin', eformatlabel)
worksheet.write(0,7,'Checkouts', eformatlabel)
worksheet.write(0,8,'Renewals', eformatlabel)
worksheet.write(0,9,'YTD', eformatlabel)
worksheet.write(0,10,'LYR', eformatlabel)
worksheet.write(0,11,'Checked Out', eformatlabel)
worksheet.write(0,12,'Barcode', eformatlabel)
worksheet.write(0,13,'Status', eformatlabel)

for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    worksheet.write(rownum+1,3,row[3], eformat)
    worksheet.write(rownum+1,4,row[4], eformat)
    worksheet.write(rownum+1,5,row[5], eformat2)
    worksheet.write(rownum+1,6,row[6], eformat2)
    worksheet.write(rownum+1,7,row[7], eformat)
    worksheet.write(rownum+1,8,row[8], eformat)
    worksheet.write(rownum+1,9,row[9], eformat)
    worksheet.write(rownum+1,10,row[10], eformat)
    worksheet.write(rownum+1,11,row[11], eformat)
    worksheet.write(rownum+1,12,row[12], eformat)
    worksheet.write(rownum+1,13,row[13], eformat)
    
    

workbook.close()


#Creating the email message
if target > 23279:
    msg = MIMEMultipart()
    msg['From'] = emailfrom
    if type(emailto) is list:
        msg['To'] = ','.join(emailto)
    else:
        msg['To'] = emailto
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = emailsubject1
    msg.attach (MIMEText(emailmessage1))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excelfile,"rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
    msg.attach(part)
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.sendmail(emailfrom, emailto, msg.as_string())

else:
    msg = MIMEMultipart()
    msg['From'] = emailfrom
    if type(emailto) is list:
        msg['To'] = ','.join(emailto)
    else:
        msg['To'] = emailto
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = emailsubject2
    msg.attach (MIMEText(emailmessage2))
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.sendmail(emailfrom, emailto, msg.as_string())
    smtp.quit()  


    








