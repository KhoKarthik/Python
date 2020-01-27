import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

username = 'karthikb9596@gmail.com'
password = '12105795'
From = username
Subject = 'OPENINGS FOR EMBEDDED AUTOMOTIVE DOMAIN'

wb = xl.load_workbook(r'C:\\Users\\ADMIN\\Desktop\\email1.xlsx')
sheet1 = wb.get_sheet_by_name('Sheet1')


names = []
emails = []

for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

#server = smtplib.SMTP('smtp.gmail.com', 587)
#server.starttls()
#server.login(username, password)

for i in range(len(emails)):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = names[i]
    msg['Subject'] = Subject
    

filename = 'Students.txt'
attachment  = open(filename,'rb')

part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition',"attachment; filename= "+filename)

format(names[i])
msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(username, password)
server.sendmail(username, emails[i], text)

#server.quit()

#server.sendmail(username, emails[i], text)
print('Mail sent to:', emails[i])

server.quit()
print('\nAll Mails are sent successfully!')
