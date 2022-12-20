###############################################################################
# DATE: 19/12/2022

# AUTHOR: Marti Rovira

# AIM: Download emails of different web email accounts into a single database.

# OUTPUT: CSV file with all emails and an html copy of all emails and attachments.

# NOTE1: I used 'Dreamhost' to purchase web domain hosts that included email use.

# NOTE2: This is the first code I share publicly. Please, feel free to reach me
# to suggest any improvements.

###############################################################################

#### Importing packages

import imaplib
import email
from email.header import decode_header
import os
import pandas as pd
import html2text
import datetime
import webbrowser
#import shutil

#### Setting date and time
today = datetime.date.today()
datetoday = today.strftime('%d/%m/%Y')
datetoday2 = today.strftime('%Y_%m_%d')

now=datetime.datetime.now()
timenow = now.strftime('%H:%M:%S')
timenow2= now.strftime('%H_%M_%S')

from datetime import datetime #

#### Setting account credentials
usernames = [
            'trial1@postmailgroup.com',
            'trial2@sppcrehabilitation.com',
            'trial3@theukmail.com'
            ]

password = '20221218_Argentina'

os.makedirs('Emails', mode = 0o777, exist_ok = True)

## Setting main database

email_db = pd.DataFrame(columns=['From', 
                                 'To', 
                                 'Date_email', 
                                 'Time_email', 
                                 'Subject', 
                                 'Text',
                                 'Date_retrieved', 
                                 'Time_retrieved', 
                                 'Saved_as'
                                 'Email',]) 


##### Iterative programme 

errors_count = 0

for username in usernames:
    try:
        # create an IMAP4 class with SSL 
        imap = imaplib.IMAP4_SSL('imap.dreamhost.com')
        # authenticate
        imap.login(username, password)

        # selecting the apropriate source folder
        status, messages = imap.select('INBOX')

        # total number of emails
        messages = int(messages[0])

        # number of top emails to fetch
        N = messages
        
        # Creating a folder to 'save' old emails.
        
        imap.create('INBOX.ProcessedEmails')

        ## Creating a database with information on the emails 

        for i in range(messages, messages-N, -1):
            # fetch the email message by ID
            res, msg2 = imap.fetch(str(i), '(RFC822)')
            for response in msg2:
                if isinstance(response, tuple):
                    # parse a bytes email into a message object
                    msg = email.message_from_bytes(response[1])
                    # decode the email subject
                    subject = decode_header(msg['Subject'])[0][0]
                    if isinstance(subject, bytes):
                        # if it's a bytes, decode to str
                        subject = subject.decode()
                    # setting email sender
                    from_ = msg.get('From')
                    to_=msg.get('To')
                    # setting email date
                    date_f = msg.get('Date')
                    len_date_f=date_f.rfind(':')+3
                    date_f=date_f[:len_date_f]
                    if date_f[0].isdigit():
                        if len(date_f)==18:
                            date_time_obj = datetime.strptime(date_f, 
                                                      '%d %b %y %H:%M:%S')
                        else:
                            date_time_obj = datetime.strptime(date_f, 
                                                      '%d %b %Y %H:%M:%S')
                    else:
                        date_time_obj = datetime.strptime(date_f, 
                                                      '%a, %d %b %Y %H:%M:%S')
                    date_d=str(datetime.date(date_time_obj))
                    # setting email time
                    date_t=datetime.time(date_time_obj)
                    date_email = date_time_obj.strftime(
                                                     '%Y_%m_%d_%H_%M_%S')

                    # setting the subject
                    subject2 = username + '_' + date_email

                    # if the email message is multipart
                    if msg.is_multipart():
                        # iterate over email parts
                        for part in msg.walk():
                            # extract content type of email
                            content_type = part.get_content_type()
                            content_disposition = str(part.get(
                                                     'Content-Disposition'))
                            try:
                                # get the email body
                                body = part.get_payload(decode=True).decode()            
                                if content_type == 'text/plain':
                                    # print only text email parts
                                    filename = f'{subject[:10]}.html'
                                    filepath = os.path.join('Emails/'+ username + '/' + subject2, filename)
                                if not os.path.isdir('Emails/' + username  + '/' + subject2):
                                    # make a folder for this email 
                                    os.makedirs('Emails/' + username + '/' + subject2, 
                                                        exist_ok=True)
                                    filepath = os.path.join('Emails/' + username + '/' + subject2, 
                                                      filename)
                                    open(filepath, 'w').write(body)
                            except:
                                pass

                            if 'attachment' in content_disposition:
                                # download attachment
                                filename = part.get_filename()
                                if filename:
                                    if not os.path.isdir(
                                         'Emails/' + username  + '/' + subject2):
                                        # make a folder for this email 
                                        os.makedirs(
                                            'Emails/' + username + '/' + subject2, 
                                            exist_ok=True)
                                    filepath = os.path.join(
                                         'Emails/' + username + '/' + subject2, 
                                          filename)
                                    # download attachment and save it
                                    open(filepath, 'wb').write(part.get_payload(
                                        decode=True))
                    else:
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == 'text/plain':
                            # print only text email parts
                            filename = f'{subject[:10]}.html'
                            filepath = os.path.join('Emails/'+ username + '/' + subject2, filename)
                            if not os.path.isdir('Emails/' + username  + '/' + subject2):
                                # make a folder for this email 
                                os.makedirs('Emails/' + username + '/' + subject2, 
                                            exist_ok=True)
                                filepath = os.path.join( 'Emails/' + username + '/' + subject2, 
                                          filename)
                            open(filepath, 'w').write(body)

                    if content_type == 'text/html':
                        # if it's HTML, create a new HTML file and open it in browser
                        if not os.path.isdir(subject):
                            # make a folder for this email (named after the subject)
                            os.makedirs(
                                 'Emails/'+ username + '/' + subject2, exist_ok=True)
                        filename = f'{subject[:10]}.html'
                        filepath = os.path.join(
                                 'Emails/'+ username + '/' + subject2, filename)
                        # write the file
                        open(filepath, 'w').write(body)
                        # open in the default browser
                        webbrowser.open(filepath)

                    body2 = html2text.html2text(body)
                    email_ = username

                    data = [(email_, 
                             from_, 
                             to_, 
                             date_d, 
                             date_t, 
                             subject, 
                             body2, 
                             datetoday, 
                             timenow, 
                             subject2
                            )]
                    data = pd.DataFrame(data)
                    data.columns = ['Email', 
                                    'From', 
                                    'To', 
                                    'Date_email', 
                                    'Time_email', 
                                    'Subject', 
                                    'Text', 
                                    'Date_retrieved', 
                                    'Time_retrieved', 
                                    'Saved_as']

                    email_db = email_db.append(data)

        ## Moving messages to 'ProcessedEmails' folder.

        imap.select('INBOX') 
        _, data = imap.uid('FETCH', '1:*' , '(RFC822.HEADER)')
        if data[0] is not None:
            messages = [data[i][0].split()[2]  for i in range(0, len(data), 2)]
            for msg_uid in messages:
                apply_lbl_msg = imap.uid('COPY', msg_uid, 'INBOX.ProcessedEmails')
                if apply_lbl_msg[0] == 'OK':
                    mov, data = imap.uid('STORE', msg_uid , '+FLAGS', '(\Deleted)')
                    imap.expunge()
                    #print('Moved msg %s' % msg_uid)
                else:
                    print(apply_lbl_msg)
                    print('Copy of msg %s did not work' % msg_uid)

        ## Closing the session

        imap.close()
        imap.logout()
        if N==0:
            print(username + ' no new emails')
        
        elif N==1:
            print(username + ': ' + str(N) + ' new email')
        
        else:
            print(username + ': ' + str(N) + ' new emails')
    
    except (ValueError, imaplib.IMAP4.error) as error :
        print(username + ': ' + 'Error')
        errors_count = errors_count + 1

email_db1 = email_db
    
#### Linking emails with profile characteristics

if len(email_db) == 0:
    
    if errors_count == 0:
        print("No callbacks this time")
        
    else:
        print("No callbacks, but errors remain")
    
    
else:

    email_db.to_csv( datetoday2 + "_" + timenow2 +'.csv', sep=';')
    
    if errors_count==0:
        print('SMS Retriever has finished with no errors!')
    
    else:
        print('SMS Retriever has finished but errors remain')

#### Finalizing emails