import imaplib
import getpass

imap_host = "outlook.office365.com"
mail_login = input("Username: ")
mail_pass = getpass.getpass()

server = imaplib.IMAP4_SSL(imap_host)
print(server.welcome)
server.login(mail_login, mail_pass)

# server = imaplib.IMAP4('outlook.office365.com') #Direccion del Servidor
# server.login('thegodzilladog@hotmail.com','*godzilla*')
# server.select()

# tipo, datos = server.search(None, 'ALL')

# server.close()
# server.logout()
