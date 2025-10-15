import smtplib
from email.message import EmailMessage
import mimetypes
from pathlib import Path


class Email:
    def __init__(self, receiver, subject, idioma,  body, files=None):
        self.sender = 'distribucio@grupserhs.com'
        self.password = 'rmbuoclwswtlvbdx'

        self.receiver = receiver.replace(';', ',')
        self.subject = subject
        self.body = body 
        self.files = files
        self.idioma = idioma

    def send(self):
        if (self.idioma=='CAT'):
            nom = 'assets/inputs/emailcatala.html'
        else:
            nom = 'assets/inputs/emailcastella.html'

        with open(nom, 'r') as arxiu:
            contingut = arxiu.read()

        msg = EmailMessage()
        msg["Subject"] = self.subject
        msg["From"] = self.sender
        msg["To"] = self.receiver
        #msg["Bcc"]= 'distribucio@grupserhs.com'
        msg.set_content(self.body)
        msg.add_alternative(self.montar_html(self.body, contingut), subtype='html')

        content_type, encoding = mimetypes.guess_type(self.files)
        if content_type is None or encoding is not None:
            content_type = "application/octet-stream"
        maintype, subtype = content_type.split("/")
        with open(self.files, "rb") as fp:
            msg.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=self.files)

        smtp = smtplib.SMTP("smtp.gmail.com", 587)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(self.sender, self.password)
        smtp.send_message(msg)
        smtp.quit()

    def montar_html(self, cos, buffer):
        html= f"{buffer}".format(cos=cos)
        return html
    