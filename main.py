import sys
from PySide6 import QtWidgets
import smtplib
import email
import imaplib
import email

class EmailReader(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Tworzenie połączenia z serwerem IMAP
        self.imap_server = 'outlook.office365.com'
        self.imap_port = 993
        self.imap_username = 'kocham.wno@outlook.com'
        self.imap_password = 'test1234test1234'

        self.imap = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
        self.imap.login(self.imap_username, self.imap_password)

        # Wybór skrzynki odbiorczej
        self.imap.select('inbox')

        # Wyszukiwanie nieprzeczytanych wiadomości
        self.status, self.response = self.imap.search(None, 'ALL')

        self.unread_msg_nums = self.response[0].split()

        # Tworzenie listy tematów nieprzeczytanych wiadomości
        self.msg_list_widget = QtWidgets.QListWidget(self)
        for num in self.unread_msg_nums:
            self.status, self.response = self.imap.fetch(num, '(RFC822)')
            self.msg = email.message_from_bytes(self.response[0][1])
            self.subject = self.msg['Subject']
            self.msg_list_widget.addItem(self.subject)

        # Przycisk do wyświetlania treści wiadomości
        self.show_msg_button = QtWidgets.QPushButton('Pokaż treść', self)
        self.show_msg_button.clicked.connect(self.show_msg)

        #self.send_mail_button = QtWidgets.QPushButton('Wyślij mail', self)
        #self.send_mail_button.clicked.connect(self.send_mail)

        # Przycisk do otwierania okna dialogowego
        self.new_mail_button = QtWidgets.QPushButton('Nowa wiadomość', self)
        self.new_mail_button.clicked.connect(self.open_mail_dialog)

        # Układanie elementów w oknie
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.msg_list_widget)
        self.layout.addWidget(self.show_msg_button)
        self.layout.addWidget(self.new_mail_button)

    def open_mail_dialog(self):
        # Tworzenie nowego okna dialogowego
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle('Nowa wiadomość')
        dialog.setModal(True)

        # Tworzenie pól tekstowych do wpisania adresata, tematu i treści wiadomości
        recipient_label = QtWidgets.QLabel('Adresat:')
        recipient_edit = QtWidgets.QLineEdit(dialog)
        subject_label = QtWidgets.QLabel('Temat:')
        subject_edit = QtWidgets.QLineEdit(dialog)
        message_label = QtWidgets.QLabel('Treść:')
        message_edit = QtWidgets.QTextEdit(dialog)

        # Układanie elementów w oknie
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.addWidget(recipient_label)
        layout.addWidget(recipient_edit)
        layout.addWidget(subject_label)
        layout.addWidget(subject_edit)
        layout.addWidget(message_label)
        layout.addWidget(message_edit)

        # Dodanie przycisków OK i Anuluj
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        # Wyświetlenie okna dialogowego i oczekiwanie na zakończenie
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            # Wysłanie wiadomości, jeśli użytkownik kliknął OK
            recipient = recipient_edit.text()
            subject = subject_edit.text()
            message = message_edit.toPlainText()
            self.send_mail(recipient, subject, message)

    def show_msg(self):
        # Odczytywanie treści wiadomości
        self.current_item = self.msg_list_widget.currentItem()
        self.current_index = self.msg_list_widget.currentRow()
        self.num = self.unread_msg_nums[self.current_index]
        self.status, self.response = self.imap.fetch(self.num, '(RFC822)')
        self.msg = email.message_from_bytes(self.response[0][1])
        for part in self.msg.walk():
            if part.get_content_type() == 'text/plain':
                self.body = part.get_payload(decode=True)
                break
        self.msg_box = QtWidgets.QMessageBox(self)
        self.msg_box.setText(self.body.decode('utf-8'))
        self.msg_box.exec()

    def send_mail(self, recipient, subject, message):
        # Dane logowania do serwera SMTP
        smtp_server = 'smtp.office365.com'
        smtp_port = 587
        smtp_username = 'kocham.wno@outlook.com'
        smtp_password = 'test1234test1234'

        # Tworzenie obiektu klasy SMTP i logowanie
        smtp = smtplib.SMTP(smtp_server, smtp_port)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(smtp_username, smtp_password)

        # Tworzenie wiadomości email
        msg = email.message.EmailMessage()
        msg['From'] = smtp_username
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.set_content(message)

        # Wysyłanie wiadomości
        smtp.send_message(msg)

        # Zamykanie połączenia SMTP
        smtp.quit()

    def closeEvent(self, event):
        # Zamykanie połączenia z serwerem IMAP
        self.imap.close()
        self.imap.logout()
        event.accept()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    email_reader = EmailReader()
    email_reader.show()
    sys.exit(app.exec())
