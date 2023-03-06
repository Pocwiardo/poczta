import smtplib
import email
import imaplib
import email

# Tworzenie połączenia z serwerem IMAP
imap_server = 'outlook.office365.com'
imap_port = 993
imap_username = 'kocham.wno@outlook.com'
imap_password = 'test1234test1234'

imap = imaplib.IMAP4_SSL(imap_server, imap_port)
imap.login(imap_username, imap_password)

# Wybór skrzynki odbiorczej
imap.select('inbox')

# Wyszukiwanie nieprzeczytanych wiadomości
status, response = imap.search(None, 'UNSEEN')

unread_msg_nums = response[0].split()

print(f"Liczba nieprzeczytanych wiadomości: {len(unread_msg_nums)}")

# Odczytywanie nagłówków nieprzeczytanych wiadomości
for num in unread_msg_nums:
    status, response = imap.fetch(num, '(RFC822)')
    msg = email.message_from_bytes(response[0][1])

    print(f"Od: {msg['From']}")
    print(f"Do: {msg['To']}")
    print(f"Temat: {msg['Subject']}")
    print(f"Data: {msg['Date']}")

    # Odczytywanie treści wiadomości
    for part in msg.walk():
        if part.get_content_type() == 'text/plain':
            body = part.get_payload(decode=True)
            print(f"Treść: {body.decode('utf-8')}")

# Zamykanie połączenia z serwerem IMAP
imap.close()
imap.logout()
