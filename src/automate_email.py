
from email.message import EmailMessage
import ssl
import smtplib
import pandas as pd

raw = pd.read_excel('/content/email.xlsx')
Data = pd.DataFrame(raw)

first_name = Data['First Name']
last_name = Data['Last Name']
company = 'ABC'
extension = 'com'

i = 0;
receivers = []
for x in last_name:
  receivers.append(first_name[i] + '.' + x + '@' + company + '.' + extension)
  i = i + 1

email_sender = 'priyansh@gmail.com'
email_passcode = '****************'

subject = "Request Regarding ** **** ***** ** at " + company + '.'

with open('Priyansh.pdf', 'rb') as resume:
  file_data = resume.read()
  file_type = 'application'
  file_name = resume.name

a = 0
while a < len(receivers):

  body = "Hi " + first_name[a] + ",\nMy name is Priyansh and I am contacting you to express my interest, and dream, and hoping for help in getting a Data or Business Analytics internship at " + company +". I am pursuing a bachelor's degree in Information and Communication Technology, and am eager to gain hands-on experience in data analysis and visualization.\nI have completed several coursework and projects that have honed my skills in data analysis and visualization using tools such as Tableau, Power BI, Excel, and Python.\nI am excited about the opportunity to work with a renowned organization like " + company + " and learn from the best in the industry.\nI would be grateful if you could consider me for the Data or Business Analytics Internship at " + company + ". I would not be the best person and I do lack in many things. But, given a chance to prove myself, I would be an asset for your company. Through each failure and rejection, I have upgraded myself and tried to do better than before. I have attached my resume for your review and I would be honored to discuss further how I can contribute.\nThanks a lot for reading my message and it means a lot to me. I look forward to hearing from you soon.\n\nBest Regards,\nPriyansh\n+918238660053" 

  email_receiver = receivers[a]
  em = EmailMessage()
  em['From'] = email_sender
  em['To'] = email_receiver
  em['Subject'] = subject
  em.set_content(body)
  em.add_attachment(file_data, maintype = 'application', subtype = file_type, filename = file_name)
  context = ssl.create_default_context()

  with smtplib.SMTP_SSL('smtp.gmail.com', 465, context = context) as smtp:
    smtp.login(email_sender, email_passcode)
    smtp.sendmail(email_sender, email_receiver, em.as_string())
  
  a = a + 1

