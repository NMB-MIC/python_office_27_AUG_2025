import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from datetime import date 


def send_to_gmail(files,username,password):

    today = date.today()
    
    send_from = 'suraphop.b@minebea.co.th'
    send_to = ['suraphop.b@minebea.co.th','devops.mic@gmail.com']
    subject = f"sale amount report on date {today}"
    text = f'''<html>
    <head>
    Dear All
    </head>
    <body>
    <div>
    I would like to report sale amount summary by daily 
    </div>
    <i> a attact file</i>
    </body>
    </html>
    '''

    file_name_1 = files.split("\\")[-1]

    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)
    msg.attach(MIMEText(text,'html'))

    #add file
    part = MIMEBase('application', "octet-stream")
    with open(files, 'rb') as f:
        file = f.read()
    part.set_payload(file)
    encoders.encode_base64(part)    
    part.add_header('Content-Disposition', f'attachment; filename="{file_name_1}"')
    msg.attach(part)

    server = smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    server.login(username,password)
    server.sendmail(send_from,send_to,msg.as_string())