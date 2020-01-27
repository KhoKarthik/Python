limport smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

username ='web mail'
password ='pswd'
From = username
Subject = 'OPENINGS FOR EMBEDDED AUTOMOTIVE DOMAIN'

wb = xl.load_workbook(r'C:\Users\ADMIN\Desktop\email1.xlsx')
sheet1 = wb.get_sheet_by_name('Sheet1')

names = []
emails = []

for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

server = smtplib.SMTP('mail.vact-tech.com', 587)
server.starttls()
server.login(username, password)

for i in range(len(emails)):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = names[i]
    msg['Subject'] = Subject
    text = '''
Dear Candidate,
                                            WELCOME TO VAct TECHNOLOGIES
    
                                                ***CONGRATULATIONS***

                        We have received your resume,for the EMBEDDED ENGINEER role,we would like to
                    invite you for walk-in interview to discuss further.


                        Based on your performance,you will be offered for inhouse trainee or an onsite
                    engineer role,which would be subject to position availability.

   

                                                  DATES FOR WALK-IN
                                                        MON-SAT
                                                  10.00 am to 5.00 pm

    Regards

    HR Dept
    VAct TECHNOLOGIES

    No 145 A2, Saradha Mill Road,
    opposite to Koushika Hospital,
    Sundarapuram, Coimbatore 641024

    Phone: +91 7871909590
    e-mail: careers@vact-tech.com
'''.format(names[i])
    msg.attach(MIMEText(text, 'plain'))
    message = msg.as_string()
    server.sendmail(username, emails[i], message)
    print('Mail sent to:', emails[i])

server.quit()
print('All Mails are sent successfully!')
