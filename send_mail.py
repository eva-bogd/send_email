import logging
import os
import smtplib
import time
from dotenv import load_dotenv
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas

load_dotenv()


logging.basicConfig(
    level=logging.INFO,
    filename='log_main.log',
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

SMTP_SERVER = 'smtp.mail.ru'
SMTP_PORT = 465
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')


def create_msg(email):
    subject = 'Тема письма'
    body = ('- Текcт письма -')
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    return msg


def format_filename(full_name):
    parts = full_name.split()
    last_name = parts[0]
    initials = parts[1][:1] + '.' + parts[2][:1] + '.'
    return f'{last_name} {initials}.pdf'


def main():
    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    server.login(EMAIL_USER, EMAIL_PASSWORD)

    df = pandas.read_excel('contacts.xlsx')

    file_count = 0
    msg_count = 0

    for col, row in df.iterrows():
        name = row['ФИО']
        email = row['Почта']

        msg = create_msg(email)

        file_name = format_filename(name)
        file_path = os.path.join('attach', file_name)
        part = MIMEBase('application', 'pdf')

        try:
            part.set_payload(open(file_path, 'rb').read())
        except FileNotFoundError:
            print(f'Не найден файл {file_name}')
            logging.error(f'Не найден файл {file_name}')
        else:
            encoders.encode_base64(part)
            part.add_header(
                            'Content-Disposition',
                            'attachment', filename=file_name,
            )
            msg.attach(part)
            file_count += 1
            logging.info(f'Прикреплен файл: {file_name}.')

        try:
            server.sendmail(EMAIL_USER, email, msg.as_string())
            msg_count += 1
            logging.info(f'Письмо отправлено на {email}, получатель: {name}.')
            time.sleep(10)
        except Exception:
            print(f'Ошибка при отправке письма на: {email}, '
                  f'получатель: {name}')
            logging.error(f'Ошибка при отправке письма на: {email}, '
                          f'получатель: {name}.')
            continue

    server.quit()

    print(f'Прикреплено файлов: {file_count}')
    print(f'Отправлено писем: {msg_count}')
    logging.info(f'SMTP-сессия завершена. Отправлено писем: {msg_count}.')


if __name__ == '__main__':
    main()
