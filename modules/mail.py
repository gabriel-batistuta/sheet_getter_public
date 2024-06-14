import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename
import os

def send_email(send_from: str, password:str,subject: str, text: str, 
send_to: list, files= None):

    # send_to= default_address if not send_to else send_to

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)  
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil: 
            ext = f.split('.')[-1:]
            attachedfile = MIMEApplication(fil.read(), _subtype = ext)
            attachedfile.add_header(
                'content-disposition', 'attachment', filename=basename(f) )
        msg.attach(attachedfile)


    smtp = smtplib.SMTP(host="smtp.gmail.com", port= 587) 
    smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

username = 'email sender'
# example of password for script app of google account 
password = 'etru zktg wqhy aheg'
default_address = ['emails receptors'] 

if __name__ == '__main__':
    planilhas = os.listdir('./')
    print(planilhas)
    send_email(send_from=username,
    password=password,         
    subject="titulo do email",
    text="texto do email",
    send_to= default_address,
    files= ['reports_lists.json', 'reports.json']
    )