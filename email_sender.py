import smtplib
import os
import time
from validate_email import validate_email

from pyfiglet import Figlet
from tqdm import tqdm
from login import SERVER, PORT, USER_EMAIL, USER_PASSWORD
import mimetypes
from email import encoders
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

def make_recipient_list():
    recipient = []
    request = 'Введите почту получателя:'

    while True:
        mail = input(request)

        if mail and validate_email(mail):
            recipient.append(mail)
            request = 'Добавьте адресата в копию или нажмите ENTER:'
            print(f'{mail} - добавлен в список получателей')
        if not mail and not recipient:
            print('Необходимо указать минимум одного получателя!')
        elif not mail:
            break
        elif not validate_email(mail):
            print(f'{mail} некорректный адрес!')

    return recipient

def validation_template():
    template = ''
    while True:
        template = input('Укажите шаблон:')

        if not template or os.path.exists(template) and os.path.splitext(template)[1] == '.html':
            break
        else:
            print('Шаблон не найден!')
    return template

def make_theme():
    theme = ''
    theme = input('Введите тему письма:')

    if not theme:
        theme = '<Без темы>'

    return theme

def make_attachments_list():
    attachments = []
    while True:
        file = input('Укажите файл для вложения или нажмите ENTER:')

        if not file:
            break

        if os.path.exists(file):
            attachments.append(file)
            print(f'{file} - добавлен')

        else:
            print('Файл не найден!')

    return attachments

def send_email(server, port, sender, password, recipient, theme, template, message, attachments):
    print('\nОтправка письма...')
    # создаем объект SMTP и передаем в параметры сервер и порт
    server = smtplib.SMTP_SSL(server, port)

    # считываем файл с шаблоном письма
    if template:
        try:
            with open(template) as file:
                template = file.read()
        except IOError:
            return 'Файл шаблона не найден'


    try:
        # логинимся
        server.login(sender, password)

        # обрабатываем кириллицу
        #msg = MIMEText(template, 'html')
        msg = MIMEMultipart()
        msg["From"] = USER_EMAIL
        msg["To"] = ','.join(recipient)
        msg["Subject"] = theme

        # Добавил текст в письмо, потом добавил шаблон в письмо
        msg.attach(MIMEText(message))
        if template:
            msg.attach(MIMEText(template, 'html'))


        if attachments:
            print('\nОбработка файлов...')
            # получил список файлов в папке attachments и сохранили
            for file in tqdm(attachments):
                print(f'\n{file}')
                # имя файла
                filename = os.path.basename(file)

                # определяем кодировку, тип файла и подтип
                ftype, encoding = mimetypes.guess_type(file)
                file_type, subtype = ftype.split('/')

                if file_type == 'text':
                    with open(file) as f:
                        file = MIMEText(f.read())
                elif file_type == 'image':
                    with open(file, 'rb') as f:
                        file = MIMEImage(f.read(), subtype)
                elif file_type == 'audio':
                    with open(file, 'rb') as f:
                        file = MIMEAudio(f.read(), subtype)
                elif file_type == 'application':
                    with open(file, 'rb') as f:
                        file = MIMEApplication(f.read(), subtype)
                else:
                    with open(file, 'rb') as f:
                        file = MIMEBase(file_type, subtype)
                        file.set_payload(f.read())
                        encoders.encode_base64(file)


            # добавляем заголовки, теперь будет прикрепляться сам файл, а не его содержимое в письмо
            file.add_header('content-disposition', 'attachment', filename=filename)
            msg.attach(file)

            time.sleep(1)

        # отправляем письмо (от кого, кому, текст сообщения)
        server.sendmail(sender, recipient, msg.as_string())
        print('Письмо отправлено!')

        return 'Письмо отправлено'

    except Exception as _ex:
        print(f'{_ex}\nНеправильный логин или пароль')

def main():
    font_text = Figlet(font='slant')
    print(font_text.renderText("Send  E-mail"))

    while True:
        recipient = make_recipient_list()
        theme = make_theme()
        template = validation_template()
        message = input('Введите текст письма:')
        attachments = make_attachments_list()

        print('ОТПРАВКА ПИСЬМА:')
        print('Получатели:', recipient)
        print('Тема:', theme)
        print('Шаблон:', template)
        print('Вложения:', attachments)

        confirmation = input('Для отправки письма нажмите ENTER, для отмены "s":')

        if not confirmation:
            break
        elif confirmation.lower() != 's':
            continue


    server = SERVER
    port = PORT
    sender = USER_EMAIL
    password = USER_PASSWORD


    send_email(server, port, sender, password, recipient, theme, template, message, attachments)


if __name__ == '__main__':
    main()