# -*- coding: utf-8 -*-
import smtplib
import os

from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase


class SendMessage(object):
    """
    Класс для отправки сообщений через почту (по-умолчанию gmail) как с вложеным .csv-файлом, так и без вложенного файла.

    Конструктор принимает логин и пароль от почты отправителя.
    """

    def __init__(self, login='', password='', smtp_server='smtp.gmail.com', port=465):
        """
        Конструктор класса для подключения к почте, по умолчанию: smtp сервер - gmail.com; порт - 465

        :param login: email отправителя
        :param password: пароль от почты отправителя
        :param smtp_server: используемый smtp сервер
        :param port: используемый порт
        """
        self._login = login
        self._password = password
        self._smtp_server = smtp_server
        self._port = port

    def _send(self, recipients, letter):
        server = smtplib.SMTP(self._smtp_server, self._port)
        try:
            server.login(self._login, self._password)
            server.sendmail(self._login, recipients, letter.as_string())
        except Exception as e:
            return e
        finally:
            server.quit()
        return 1

    def send_message(self, recipients=(), subject='', text_message='', attach_file=None):
        """
        Метод для отправки сообщения.

        :param recipients: Строка с адресом получателя, если получатель один.
        Список с адресами, если получателей несколько.
        :param subject: Строка, тема письма
        :param text_message: Строка, тело письма
        :param attach_file: Строка, путь до файла вложения
        :return:
        Возвращает 1 - если сообщение удачно отправлено, впротивном случае описание ошибки
        """
        if len(recipients) == 0:
            raise Exception('Нет адресата для отправки')

        letter = MIMEMultipart()
        letter["From"] = self._login
        letter["Subject"] = subject
        letter["Content-Type"] = 'text/html; charset=utf-8'
        # letter["To"] = ", ".join(recipients)
        letter.attach(MIMEText(text_message, 'plain'))

        if attach_file:
            attachement = MIMEBase('application', 'csv')

            file_name = os.path.basename(attach_file)
            try:
                attachement.set_payload(open(attach_file, 'rb').read())
                attachement.add_header('Content-Disposition',
                                       f'attachement; filename={file_name}')
                encoders.encode_base64(attachement)
                letter.attach(attachement)
            except:
                letter.attach(MIMEText('\n\nВложение не прикреплено, проверьте путь до вложения', 'plain'))

        return self._send(recipients, letter)
