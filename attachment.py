#extract attachment

import imaplib
import email
import sys
import os

def extrct_attachment(imap_server, user, pwd):
    try :
        #getting the mail 
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(user, pwd)
        mail.select("inbox")

        result, data = mail.search(None, "subject keyword") 
        ids_list = data[0].split()
        latest_email_id = ids_list[-1]

        result, data = mail.fetch(latest_email_id, "(RFC822)")
        raw_mail =  email.message_from_bytes(data[0][1])

        #process xlsx attachment

        #load xlsx attachment as data frame

    except Exception as e : 
        print(f'an error has occurred : {e}')
        return None
    
    finally :
        mail.logout()
        mail.close()





