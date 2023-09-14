import mimetypes
import os

from docx2pdf import convert
from docxtpl import DocxTemplate

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase


def docx_reader(filename):
    doc = DocxTemplate(filename)
    with open("C:\\Users\\Timofei\\Desktop\\csv\\Drept_send.csv", "r", encoding="utf-8") as file:
        with open("C:\\Users\\Timofei\\Desktop\\csv\\Error.txt", "w+", encoding="utf-8") as errorFile:
            file_data = file.read()
            lines = file_data.split("\n")
            for line in lines:
                fields = line.split(",")
                email = fields[4]
                password = fields[5]
                full_name = fields[0] + " " + fields[1]
                send_email = fields[6]
                context = {'email': email, 'password': password}
                doc.render(context)
                doc.save("C:\\Users\\Timofei\\Desktop\\word_files/" + full_name + ".docx")
                convert("C:\\Users\\Timofei\\Desktop\\word_files/" + full_name + ".docx",
                        "C:\\Users\\Timofei\\Desktop\\pdf_files/" + full_name + ".pdf")
                os.remove("C:\\Users\\Timofei\\Desktop\\word_files/" + full_name + ".docx")
                try:
                    send_message(send_email, "C:\\Users\\Timofei\\Desktop\\pdf_files/" + full_name + ".pdf")
                    print("successfully sent email to - " + send_email + " NP - " + full_name)
                except:
                    print("Error sending message to address " + send_email + " NP - " + full_name)
                    errorFile.write(line + '\n')


def send_message(email_address, file_address):
    password = "Wwq807627"

    msg = MIMEMultipart()
    msg['From'] = "info.5@usm.md"
    msg['To'] = email_address
    msg['Subject'] = "Cont Microsoft 365"

    # body = "Sample text"
    # msg.attach(MIMEText(body, 'plain'))

    html = """
        <!DOCTYPE html>
        <html>
            <head>
                <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
            </head>
            <body>
                <p><font color="red" face="Arial">LIMITARE DE OBLIGATIUNI: Acest e-mail sau atasament contine informatii care pot fi, partial sau in intregime, protejate de lege. Orice utilizare sau transmitere neautorizata a acestui mesaj, totala sau partiala, este strict interzisa. Aceste informatii sunt adresate doar destinatarului si pot sa nu exprime punctele de vedere ale Universitatii de Stat din Moldova. In cazul in care o eroare de transmitere a directionat gresit acest e-mail, va rugam sa notificati autorul printr-un raspuns la mesaj. Daca nu sunteti destinatarul vizat, nu aveti permisiunea sa dezvaluiti, sa distribuiti, sa copiati, sa tipariti sau sa utilizati acest e-mail.</font></p>
            </body>
        </html>"""
    msg.attach(MIMEText(html, 'html', 'utf-8'))

    filepath = file_address
    filename = os.path.basename(filepath)

    if os.path.isfile(filepath):
        ctype, encoding = mimetypes.guess_type(filepath)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        if maintype == 'text':
            with open(filepath) as fp:
                file = MIMEText(fp.read(), _subtype=subtype)
                fp.close()
        elif maintype == 'image':
            with open(filepath, 'rb') as fp:
                file = MIMEImage(fp.read(), _subtype=subtype)
                fp.close()
        elif maintype == 'audio':
            with open(filepath, 'rb') as fp:
                file = MIMEAudio(fp.read(), _subtype=subtype)
                fp.close()
        else:
            with open(filepath, 'rb') as fp:
                file = MIMEBase(maintype, subtype)
                file.set_payload(fp.read())
                fp.close()
            encoders.encode_base64(file)
        file.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(file)

    server = smtplib.SMTP(host='smtp.office365.com', port=587)
    server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()


def main():
    docx_reader("shablon.docx")


main()
