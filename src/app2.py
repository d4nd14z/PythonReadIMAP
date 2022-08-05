import imaplib
import email

# adapted from: https://stackoverflow.com/a/53827775/2327328

def read_email_from_gmail():
	mail = imaplib.IMAP4_SSL('outlook.office365.com')
	mail.login('thegodzilladog@hotmail.com','*godzilla*')
	mail.select('inbox')

	result, data = mail.search(None, '(UNSEEN)')
	mail_ids = data[0]

	id_list = mail_ids.split()   

	for _, i in enumerate(id_list):
		# need str(int(i))
		result, data = mail.fetch(str(int(i)), '(RFC822)' )
		for response_part in data:
			if isinstance(response_part, tuple):
				# from_bytes, not from_string
				msg = email.message_from_bytes(response_part[1])
				email_subject = msg['subject']
				email_from = msg['from']
				print ('From : ' + email_from + '\n')
				print ('Subject : ' + email_subject + '\n')

# nothing to print here
read_email_from_gmail()