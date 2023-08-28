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


def docxreader(filename):
    doc = DocxTemplate(filename)
    with open("C:\\Users\\Timofei\\Desktop\\csv\\dplogistica.csv", "r", encoding="utf-8") as file:
        file_data = file.read()
        lines = file_data.split("\n")
        for line in lines:
            fields = line.split(",")
            email = fields[4]
            password = fields[5]
            full_name = fields[0] + " " + fields[1]
            context = {'email': email, 'password': password}
            doc.render(context)
            doc.save("C:\\Users\\Timofei\\Desktop\\word_files/" + full_name + ".docx")
            convert("C:\\Users\\Timofei\\Desktop\\word_files/" + full_name + ".docx",
                    "C:\\Users\\Timofei\\Desktop\\pdf_files/" + full_name + ".pdf")
            os.remove("C:\\Users\\Timofei\\Desktop\\word_files/" + full_name + ".docx")
            #sendmessage( , "C:\\Users\\Timofei\\Desktop\\pdf_files/" + full_name + ".pdf")


def sendmessage(email_address, file_address):

    password = "Wwq807627"

    msg = MIMEMultipart()
    msg['From'] = "info.5@usm.md"
    msg['To'] = email_address
    msg['Subject'] = "Cont Microsoft 365"

    body = "Текст сообщения"
    msg.attach(MIMEText(body, 'plain'))

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

    filepath = file_address  # Имя файла в абсолютном или относительном формате
    filename = os.path.basename(filepath)  # Только имя файла

    if os.path.isfile(filepath):  # Если файл существует
        ctype, encoding = mimetypes.guess_type(filepath)  # Определяем тип файла на основе его расширения
        if ctype is None or encoding is not None:  # Если тип файла не определяется
            ctype = 'application/octet-stream'  # Будем использовать общий тип
        maintype, subtype = ctype.split('/', 1)  # Получаем тип и подтип
        if maintype == 'text':  # Если текстовый файл
            with open(filepath) as fp:  # Открываем файл для чтения
                file = MIMEText(fp.read(), _subtype=subtype)  # Используем тип MIMEText
                fp.close()  # После использования файл обязательно нужно закрыть
        elif maintype == 'image':  # Если изображение
            with open(filepath, 'rb') as fp:
                file = MIMEImage(fp.read(), _subtype=subtype)
                fp.close()
        elif maintype == 'audio':  # Если аудио
            with open(filepath, 'rb') as fp:
                file = MIMEAudio(fp.read(), _subtype=subtype)
                fp.close()
        else:  # Неизвестный тип файла
            with open(filepath, 'rb') as fp:
                file = MIMEBase(maintype, subtype)  # Используем общий MIME-тип
                file.set_payload(fp.read())  # Добавляем содержимое общего типа (полезную нагрузку)
                fp.close()
            encoders.encode_base64(file)  # Содержимое должно кодироваться как Base64
        file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
        msg.attach(file)

    server = smtplib.SMTP(host='smtp.office365.com', port=587)
    server.starttls()
    server.login(msg['From'], password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    print("successfully sent email to %s:" % (msg['To']))
    server.quit()


def main():
    docxreader("shablon.docx")


main()
