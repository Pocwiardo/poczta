import sys
from PySide6 import QtWidgets, QtCore
from PySide6.QtGui import QFont, QFontMetrics
import smtplib
import imaplib
import email
import os
from gensim.models import KeyedVectors

class EmailReader(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.setGeometry(100, 100, 640, 480)
        self.imap_server = 'outlook.office365.com'
        self.imap_port = 993
        self.imap_username = 'kocham.wno@outlook.com'
        self.imap_password = 'test1234test1234'

        self.imap = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
        self.imap.login(self.imap_username, self.imap_password)

        self.imap.select('inbox')


        #self.model = KeyedVectors.load_word2vec_format('nkjp+wiki-forms-restricted-100-skipg-ns.txt')
        self.model = KeyedVectors.load_word2vec_format('GoogleNews-vectors-negative300.bin', binary=True, limit=20000)
        self.unread_font = QFont()
        self.unread_font.setBold(True)
        self.status, self.response = self.imap.search(None, 'ALL')

        self.msg_nums = self.response[0].split()
        self.unseen_msg_nums = self.imap.search(None, 'UNSEEN')[1][0].split()
        #self.auto_reply()

        self.msg_list_widget = QtWidgets.QListWidget(self)
        for num in self.msg_nums:
            self.status, self.response = self.imap.fetch(num, '(RFC822)')
            self.msg = email.message_from_bytes(self.response[0][1])
            self.subject = self.msg['Subject']
            self.msg_list_widget.addItem(self.subject)

        self.show_msg_button = QtWidgets.QPushButton('Pokaż treść', self)
        self.show_msg_button.clicked.connect(self.show_msg)

        #self.send_mail_button = QtWidgets.QPushButton('Wyślij mail', self)
        #self.send_mail_button.clicked.connect(self.send_mail)

        self.new_mail_button = QtWidgets.QPushButton('Nowa wiadomość', self)
        self.new_mail_button.clicked.connect(self.open_mail_dialog)

        self.refresh_button = QtWidgets.QPushButton('Odśwież', self)
        self.refresh_button.clicked.connect(self.refresh_messages)


        self.keyword_textbox = QtWidgets.QLineEdit(self)
        self.keyword_textbox.setPlaceholderText('Wpisz słowo kluczowe')
        self.keyword_textbox.returnPressed.connect(self.refresh_messages)


        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.keyword_textbox)
        self.layout.addWidget(self.msg_list_widget)
        self.layout.addWidget(self.show_msg_button)

        self.layout.addWidget(self.new_mail_button)
        self.layout.addWidget(self.refresh_button)
        self.refresh_messages()
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.refresh_messages)
        self.timer.start(30000)

    def is_matching_message(self, text, keyword):

        if keyword.lower() in text.lower():
            return True
        for word in text.lower().split():
            for keywords in keyword.lower().split():
                try:
                    similarity = self.model.similarity(word, keywords)

                except:
                    similarity = 0
                if similarity > 0.4:
                    #print(keywords)
                    #print(word)
                    #print(similarity)
                    return True
        return False

    def refresh_messages(self):
        self.imap.select('inbox')
        self.status, self.response = self.imap.search(None, 'ALL')
        self.msg_nums = self.response[0].split()
        self.unseen_msg_nums = self.imap.search(None, 'UNSEEN')[1][0].split()

        self.msg_list_widget.clear()

        keyword = self.keyword_textbox.text().lower().strip()

        for num in self.msg_nums:
            self.status, self.response = self.imap.fetch(num, '(RFC822)')
            self.msg = email.message_from_bytes(self.response[0][1])
            self.subject = self.msg['Subject']
            font = QFont()
            if num in self.unseen_msg_nums:
                font.setBold(True)

            if keyword != '':
                body = ''
                for part in self.msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload(decode=True).decode('utf-8')
                        break
                body = body + self.subject
                if self.is_matching_message(body, keyword):
                    item = QtWidgets.QListWidgetItem(self.subject)
                    item.setFont(font)
                    item.setData(QtCore.Qt.UserRole, num)
                    self.msg_list_widget.addItem(item)
            else:
                item = QtWidgets.QListWidgetItem(self.subject)
                item.setFont(font)
                item.setData(QtCore.Qt.UserRole, num)
                self.msg_list_widget.addItem(item)

        self.auto_reply()

    def auto_reply(self):
        #self.imap.select('inbox')
        #self.status, self.response = self.imap.search(None, 'UNSEEN')

        for num in self.unseen_msg_nums:
            self.status, self.response = self.imap.fetch(num, '(RFC822)')
            email_data = self.response[0][1]
            message = email.message_from_bytes(email_data)

            sender = message['From']
            subject = message['Subject']

            subject = 'Re: ' + subject
            replied_message = 'Automatyczna odpowiedź: Otrzymałem twoją wiadomość. Odpiszę na nią tak szybko jak to możliwe'

            self.send_mail(sender,subject,replied_message)

        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.auto_reply)
        self.timer.start(60000)

    def open_mail_dialog(self):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle('Nowa wiadomość')
        dialog.setModal(True)

        recipient_label = QtWidgets.QLabel('Adresat:')
        recipient_edit = QtWidgets.QLineEdit(dialog)
        subject_label = QtWidgets.QLabel('Temat:')
        subject_edit = QtWidgets.QLineEdit(dialog)
        message_label = QtWidgets.QLabel('Treść:')
        message_edit = QtWidgets.QTextEdit(dialog)

        layout = QtWidgets.QVBoxLayout(dialog)
        layout.addWidget(recipient_label)
        layout.addWidget(recipient_edit)
        layout.addWidget(subject_label)
        layout.addWidget(subject_edit)
        layout.addWidget(message_label)
        layout.addWidget(message_edit)

        attach_button = QtWidgets.QPushButton('Załącz plik', dialog)
        layout.addWidget(attach_button)

        def attach_file():
            file_path, _ = QtWidgets.QFileDialog.getOpenFileName(dialog, 'Wybierz plik do załączenia', '',
                                                                 'Pliki (*.*)')
            if file_path:
                attachments.append(file_path)

        attach_button.clicked.connect(attach_file)

        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        attachments = []
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            recipient = recipient_edit.text()
            subject = subject_edit.text()
            message = message_edit.toPlainText()
            self.send_mail(recipient, subject, message, attachments)

    def send_mail(self, recipient, subject, message, attachments=None):
        smtp_server = 'smtp.office365.com'
        smtp_port = 587
        smtp_username = 'kocham.wno@outlook.com'
        smtp_password = 'test1234test1234'

        smtp = smtplib.SMTP(smtp_server, smtp_port)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(smtp_username, smtp_password)

        msg = email.message.EmailMessage()
        msg['From'] = smtp_username
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.set_content(message)

        if attachments:
            for attachment_path in attachments:
                with open(attachment_path, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment_path)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        smtp.send_message(msg)

        smtp.quit()

    def show_msg(self):
        self.current_item = self.msg_list_widget.currentItem()
        self.current_index = self.msg_list_widget.currentRow()
        self.num = self.current_item.data(QtCore.Qt.UserRole)
        #self.unseen_msg_nums = self.imap.search(None, 'UNSEEN')[1][0].split()
        if self.num in self.unseen_msg_nums:
            #print("Nieprzeczytana")
            self.status, self.response = self.imap.fetch(self.num, '(RFC822)')
            email_data = self.response[0][1]
            message = email.message_from_bytes(email_data)

            sender = message['From']
            subject = message['Subject']

            subject = 'Re: ' + subject
            replied_message = 'Automatyczna odpowiedź: Odczytałem twoją wiadomość. '

            self.send_mail(sender, subject, replied_message)
        #else:
        #    print("przeczytana")
        self.status, self.response = self.imap.fetch(self.num, '(RFC822)')
        self.msg = email.message_from_bytes(self.response[0][1])
        attachments = []
        for part in self.msg.walk():
            if part.get_content_disposition() == 'attachment':
                attachments.append(part)
        body = ''
        for part in self.msg.walk():
            if part.get_content_type() == 'text/plain':
                body = part.get_payload(decode=True).decode('utf-8')
                break

        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(self.msg['Subject'])
        dialog.setModal(True)
        layout = QtWidgets.QVBoxLayout(dialog)

        message_label = QtWidgets.QLabel('Treść:')
        message_edit = QtWidgets.QTextEdit(dialog)
        message_edit.setReadOnly(True)
        message_edit.setText(body)
        layout.addWidget(message_label)
        layout.addWidget(message_edit)

        if attachments:
            attachments_label = QtWidgets.QLabel('Załączniki:')
            attachments_list = QtWidgets.QListWidget(dialog)
            layout.addWidget(attachments_label)
            layout.addWidget(attachments_list)
            for attachment in attachments:
                attachments_list.addItem(attachment.get_filename())

            download_button = QtWidgets.QPushButton('Pobierz załącznik', dialog)
            layout.addWidget(download_button)
            download_button.clicked.connect(
                lambda: self.download_attachment(dialog, attachments_list.currentItem().text(), attachments))

        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)


        dialog.exec()
        self.refresh_messages()

    def download_attachment(self, dialog, filename, attachments):
        for attachment in attachments:
            if attachment.get_filename() == filename:
                data = attachment.get_payload(decode=True)
                with open(filename, 'wb') as f:
                    f.write(data)
                QtWidgets.QMessageBox.information(dialog, 'Pobieranie zakończone',
                                                  f'Załącznik {filename} został pobrany.')
                break


    def closeEvent(self, event):
        self.imap.close()
        self.imap.logout()
        event.accept()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    email_reader = EmailReader()
    email_reader.show()
    sys.exit(app.exec())
