#mass emailer for the hackathon
#Patrick Riley, inspired by http://yuqli.com/?p=653
#12/1/17
#Python 3.6.1 (v3.6.1:69c0db5, Mar 21 2017, 18:41:36)


import openpyxl, pprint, smtplib, base64
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.message import EmailMessage


# Read tables into a Python object
print ('Opening workbook...')
wb = openpyxl.load_workbook('School List.xlsx')
# Get all sheet names
wb.get_sheet_names()

#gets the specific sheet
response = wb.get_sheet_by_name('Assignments') # Note here name is case sensitive
#response = wb.get_sheet_by_name('Patrick')
print(response)

#used to save the corresponding names and emails
names = []
emails = []
print ('Reading rows...')

#reads in all the rows of names and emails and then adds them to their respective lists

for row in range(39, 72): # for loop is open in the end
    names.append(str(response['E'+str(row)].value).strip()) # Note here how to manipulate strings in Python
    emails.append(str(response['F'+str(row)].value).strip())

#names.append(str(response['C13'].value).strip())
#emails.append(str(response['D13'].value).strip())

#email server stuff, login and connect
smtpObj = smtplib.SMTP('smtp.gmail.com:587')
smtpObj.starttls() # Upgrade the connection to a secure one using TLS (587)
# If using SSL encryption (465), can skip this step
smtpObj.ehlo() # To start the connection
smtpObj.login('hackbi@bishopireton.org', 'Windows12')


#this is where the email actually is created and sent
for i in range(0, len(emails)):

    #creates an email and adds some basic attributes
    msg = EmailMessage()
    msg['From'] = 'HackBI <hackbi@bishopireton.org>' # Note the format
    msg['To'] = '%s <' % names[i] +emails[i] + '>'
    msg['Subject'] = 'Bishop Ireton High School Hackathon!'

    #reads in the text for the email body from a text file and sets it to the email
    with open("form letter.txt") as letter:
        message = letter.read() % names[i]    
    msg.set_content(message)

    #print("attaching pdf")

    #reads in a file and adds it as an attachment to the email
    with open("HackBI 2018 Flyer.pdf","rb") as f:
        file = f.read()    
    msg.add_attachment(file, filename="HackBI 2018 Flyer",maintype='application', subtype='pdf')

    #print("sending email(s)")
    #sends the email to that email being iterated through
    smtpObj.sendmail('hackbi@bishopireton.org',emails[i],msg.as_string())
    print("Email sent to " + str(emails[i]))

#closes everything up
smtpObj.quit()

