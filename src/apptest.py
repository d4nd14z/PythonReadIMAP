import imaplib
import email
from email.header import decode_header
import webbrowser
import os
from getpass import getpass


# Datos del usuario
username = input("Correo: ")
password = getpass("Password: ")


# Crear conexion
imap = imaplib.IMAP4_SSL("outlook.office365.com")

#Iniciar sesion
imap.login(username, password)

#status, mensajes = imap.select("INBOX") #Leer todos los mensajes de la bandeja de entrada
status, mensajes = imap.select("UNSEEN") #Leer todos los mensajes de la bandeja de entrada

#Retorna[b'1'] => Indica que hay 1 correo pendiente por leer

unreadMessages = int(mensajes[0])

for i in range(unreadMessages, 0, -1):
    try: 
        #Intentar leer el mail, en caso que falle sale del ciclo
        res, mensaje = imap.fetch(str(i), "(RFC822)") #Estandar RFC822: Obtener todas las partes que componen el correo
    except: 
        break

    for respuesta in mensaje:
        if isinstance(respuesta, tuple):
            mensaje = email.message_from_bytes(respuesta[1])
            #Decodificar el encabezado del mensaje
            subject = decode_header(mensaje["Subject"])[0][0]  #Obtener el asunto del mensaje
            if isinstance(subject, bytes):
                subject = subject.decode()
            msgFrom = mensaje.get("From") #Obtener quien envia el mensaje
            
            print("Asunto: ", subject)
            print("From: ", msgFrom)
            
            #Si el mensaje es html se procesa de forma distinta
            if mensaje.is_multipart():
                for part in mensaje.walk(): #Obtiene cada una de las partes del correo
                    #Extraer el contenido del mensaje
                    contentType = part.get_content_type()
                    contentDisposition = str(part.get("Content-Disposition"))
                    try: 
                        #Intentar leer el cuerpo del correo
                        body = part.get_payload(decode=True).decode()
                    except: 
                        pass

                    #Si el contenido del mensaje es texto plano y no tiene un archivo adjunto...
                    if contentType == "text/plain" and "attachment" not in contentDisposition:
                        print(body)
                    elif "attachment" in contentDisposition:
                        fileName = part.get_filename()
                        #Si si se encontraron archivos adjuntos
                        if fileName: 
                            #Crear el directorio "downloads" en caso de que no exista
                            if not os.path.isdir("downloads"):                                
                                os.mkdir(os.path.join("downloads"))
                            #Descargar los adjuntos y guardar los archivos en el directorio "downloads"     
                            pathFile = os.path.join("downloads", fileName)
                            open(pathFile, "wb").write(part.get_payload(decode=True))

                            ###################################################
                            # Luego de guardar el archivo, debe leerse el pdf
                            ###################################################



imap.close()  #Cerrar conexion IMAP
imap.logout() #Cerrar correo electronico


