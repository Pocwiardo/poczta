import sys
from PySide6 import QtWidgets, QtCore
import smtplib
import email
import imaplib
import email
import os
from gensim.models import KeyedVectors

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

        self.auto_reply()
        #self.model = KeyedVectors.load_word2vec_format('nkjp+wiki-forms-restricted-100-skipg-ns.txt')
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

        # Przycisk odświeżania
        self.refresh_button = QtWidgets.QPushButton('Odśwież', self)
        self.refresh_button.clicked.connect(self.refresh_messages)

        # Pole tekstowe dla słowa kluczowego
        self.keyword_textbox = QtWidgets.QLineEdit(self)
        self.keyword_textbox.setPlaceholderText('Wpisz słowo kluczowe')
        self.keyword_textbox.returnPressed.connect(self.refresh_messages)

        # Układanie elementów w oknie
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.keyword_textbox)
        self.layout.addWidget(self.msg_list_widget)
        self.layout.addWidget(self.show_msg_button)

        self.layout.addWidget(self.new_mail_button)
        self.layout.addWidget(self.refresh_button)
        self.refresh_messages()


    def is_matching_message(self, text, keyword):
        """
        Sprawdza, czy dana wiadomość zawiera słowo klucz lub słowo do niego podobne.
        """
        # Sprawdzanie tematu wiadomości
        if keyword in text.lower():
            return True
            #try:
            #    similarity = self.model.similarity(w, keyword)
            #except:
            #    similarity = 0
            #if similarity > 0.7:
            #    return True
        return False

    def refresh_messages(self):
        # Wyszukiwanie nieprzeczytanych wiadomości
        self.imap.select('inbox')
        self.status, self.response = self.imap.search(None, 'ALL')
        self.unread_msg_nums = self.response[0].split()

        # Czyszczenie listy tematów nieprzeczytanych wiadomości
        self.msg_list_widget.clear()

        keyword = self.keyword_textbox.text().lower().strip()

        # Tworzenie listy tematów nieprzeczytanych wiadomości
        for num in self.unread_msg_nums:
            self.status, self.response = self.imap.fetch(num, '(RFC822)')
            self.msg = email.message_from_bytes(self.response[0][1])
            self.subject = self.msg['Subject']

            if keyword != '':
                body = ''
                for part in self.msg.walk():
                    if part.get_content_type() == 'text/plain':
                        body = part.get_payload(decode=True).decode('utf-8')

                        break
                body = body + self.subject
                if self.is_matching_message(body, keyword):
                    item = QtWidgets.QListWidgetItem(self.subject)
                    item.setData(QtCore.Qt.UserRole, num)
                    self.msg_list_widget.addItem(item)
            else:
                item = QtWidgets.QListWidgetItem(self.subject)
                item.setData(QtCore.Qt.UserRole, num)
                self.msg_list_widget.addItem(item)

    def auto_reply(self):
        # Wyszukiwanie nieprzeczytanych wiadomości
        self.imap.select('inbox')
        self.status, self.response = self.imap.search(None, 'UNSEEN')

        for num in self.response[0].split():
            # Pobieranie nieprzeczytanych wiadomości
            self.status, self.response = self.imap.fetch(num, '(RFC822)')
            email_data = self.response[0][1]
            message = email.message_from_bytes(email_data)

            # Pobieranie nadawcy i tematu wiadomości
            sender = message['From']
            subject = message['Subject']

            # Tworzenie odpowiedzi na wiadomość
            #reply = email.message.EmailMessage()
            #reply['To'] = sender
            #reply['From'] = self.imap_username
            #reply['Subject'] = 'Re: ' + subject
            subject = 'Re: ' + subject
            #reply.set_content('Automatyczna odpowiedź: Otrzymałem twoją wiadomość.')
            replied_message = 'Automatyczna odpowiedź: Otrzymałem twoją wiadomość. Odpiszę na nią tak szybko jak to możliwe'

            # Wysyłanie odpowiedzi
            #self.smtp.send_message(reply)
            self.send_mail(sender,subject,replied_message)

        # Oczekiwanie na kolejne wiadomości
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.auto_reply)
        self.timer.start(60000)  # Czas w milisekundach (1 minuta = 60 000 ms)

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

        # Dodanie przycisku do wyboru plików do załączenia
        attach_button = QtWidgets.QPushButton('Załącz plik', dialog)
        layout.addWidget(attach_button)

        # Funkcja do obsługi wyboru pliku
        def attach_file():
            file_path, _ = QtWidgets.QFileDialog.getOpenFileName(dialog, 'Wybierz plik do załączenia', '',
                                                                 'Pliki (*.*)')
            if file_path:
                # Dodanie pliku do listy załączników
                attachments.append(file_path)

        # Podpięcie funkcji attach_file do przycisku
        attach_button.clicked.connect(attach_file)

        # Dodanie przycisków OK i Anuluj
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        # Wyświetlenie okna dialogowego i oczekiwanie na zakończenie
        attachments = []
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            # Wysłanie wiadomości, jeśli użytkownik kliknął OK
            recipient = recipient_edit.text()
            subject = subject_edit.text()
            message = message_edit.toPlainText()
            self.send_mail(recipient, subject, message, attachments)

    def send_mail(self, recipient, subject, message, attachments=None):
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

        # Create the email message object
        msg = email.message.EmailMessage()
        msg['From'] = smtp_username
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.set_content(message)

        # Add attachments to the message
        if attachments:
            for attachment_path in attachments:
                with open(attachment_path, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment_path)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        # Send the message
        smtp.send_message(msg)

        # Close the SMTP connection
        smtp.quit()

    def show_msg(self):
        # Odczytywanie treści wiadomości
        self.current_item = self.msg_list_widget.currentItem()
        self.current_index = self.msg_list_widget.currentRow()
        self.num = self.current_item.data(QtCore.Qt.UserRole)
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

        # Tworzenie nowego okna z treścią wiadomości i listą załączników
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle(self.msg['Subject'])
        dialog.setModal(True)
        layout = QtWidgets.QVBoxLayout(dialog)

        # Dodanie pola tekstowego z treścią wiadomości
        message_label = QtWidgets.QLabel('Treść:')
        message_edit = QtWidgets.QTextEdit(dialog)
        message_edit.setReadOnly(True)
        message_edit.setText(body)
        layout.addWidget(message_label)
        layout.addWidget(message_edit)

        # Dodanie listy załączników
        if attachments:
            attachments_label = QtWidgets.QLabel('Załączniki:')
            attachments_list = QtWidgets.QListWidget(dialog)
            layout.addWidget(attachments_label)
            layout.addWidget(attachments_list)
            for attachment in attachments:
                attachments_list.addItem(attachment.get_filename())

            # Dodanie przycisku pobierania załącznika
            download_button = QtWidgets.QPushButton('Pobierz załącznik', dialog)
            layout.addWidget(download_button)
            download_button.clicked.connect(
                lambda: self.download_attachment(dialog, attachments_list.currentItem().text(), attachments))

        # Dodanie przycisku OK
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        # Wyświetlenie okna dialogowego i oczekiwanie na zakończenie
        dialog.exec()

    def download_attachment(self, dialog, filename, attachments):
        # Pobieranie wybranego załącznika i zapisywanie do pliku
        for attachment in attachments:
            if attachment.get_filename() == filename:
                data = attachment.get_payload(decode=True)
                with open(filename, 'wb') as f:
                    f.write(data)
                QtWidgets.QMessageBox.information(dialog, 'Pobieranie zakończone',
                                                  f'Załącznik {filename} został pobrany.')
                break


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
