import email
import smtplib
import os
import mimetypes
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

addressFrom: str = os.environ['MAIL_USER']
password: str = os.environ['MAIL_PASSWORD']

def send_email(addressTo: str, msgSubject: str, msgText: str, file: str) -> None:
    msg: email.mime.multipart.MIMEMultipart = MIMEMultipart()
    msg['From']: str = addressFrom
    msg['To']: str = addressTo
    msg['Subject']: str = msgSubject
    body: str = msgText
    msg.attach(MIMEText(body, 'plain'))

    process_attachement(msg, file)

    server: smtplib.SMTP_SSL = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(addressFrom, password)
    server.send_message(msg)
    server.quit()


def process_attachement(msg: email.mime.multipart.MIMEMultipart, file: str) -> None:
    if os.path.isfile(file):
        attach_file(msg, file)


def attach_file(msg: email.mime.multipart.MIMEMultipart, filepath: str) -> None:
    filename: str = os.path.basename(filepath)
    ctype: str = mimetypes.guess_type(filepath)[0]
    encoding: str = mimetypes.guess_type(filepath)[1]
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype: str = ctype.split('/', 1)[0]
    subtype: str = ctype.split('/', 1)[1]
    if maintype == 'text':
        with open(filepath) as fp:
            file: email.mime.text = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
    else:
        with open(filepath, 'rb') as fp:
            file: email.mime.base = MIMEBase(maintype, subtype)
            file.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(file)
