# Automated Email Delivery System in Python
# By : Shekhar Saxena
#-------------------------------------------------


import smtplib #For connecting to Email Servers
from email.message import EmailMessage #Email Module in Python
from string import Template #To Customize the HTML file as our need
from pathlib import Path #To connect html with python file
import pandas as pd #To extract data from Excel File

df = pd.read_excel('data.xlsx',sheet_name=0) 
#sheet_name is index of sheet because excel contails multiple sheets.
#start with 1 to avoid zeroth row which contains Column Names.
name = list(df['Names']) #Replace with column names in excel file
marks = list(df['Marks'])
mails = list(df['Emails'])
sec = list(df['Sec'])

for i in range(0,4): #Range should be 1 to your last row

	mark = Template(Path('index.html').read_text())
	email = EmailMessage()
	email['from'] = 'Shekhar Saxena' #Name of Sender
	email['to'] = mails[i] 
	email['subject'] = 'Email Testing with Python' #Subject of mail

	#Customize HTML file using Template
	email.set_content(mark.substitute(name= name[i],mark= marks[i],section = sec[i]),subtype='html')
	with smtplib.SMTP(host='smtp.gmail.com', port = 587) as smtp:
		#host and port are different for different Email Services. Here we are using Gmail.
	    smtp.ehlo() 
	    smtp.starttls()
	    smtp.login('your_gmail', 'your_password')
	    smtp.send_message(email)

print('Mail Sent.')
