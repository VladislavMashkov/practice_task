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
    msg = MIMEMultipart()
    msg['From'] = addressFrom
    msg['To'] = addressTo
    msg['Subject'] = msgSubject
    body = msgText
    msg.attach(MIMEText(body, 'plain'))

    process_attachement(msg, file)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(addressFrom, password)
    server.send_message(msg)
    server.quit()


def process_attachement(msg, file: str) -> None:
    if os.path.isfile(file):
        attach_file(msg, file)


def attach_file(msg, filepath) -> None:
    filename = os.path.basename(filepath)
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    if maintype == 'text':
        with open(filepath) as fp:
            file = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
    else:
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)
            file.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(file)