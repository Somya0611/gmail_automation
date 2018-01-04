from xlrd import open_workbook
import email
import getpass, imaplib
import os
import sys


def excel_sheet():
    wb = open_workbook('input_data.xls')
    for s in wb.sheets():
        #print 'Sheet:',s.name
        values = []
        for row in range(s.nrows):
            col_value = []
            for col in range(s.ncols):
                value  = (s.cell(row,col).value)
                try : value = str(int(value))
                except : pass
                col_value.append(value)
            values.append(col_value)
    print values
    wb.close()
    
def login_download_attachment():
    detach_dir = '.'
    if 'attachments' not in os.listdir(detach_dir):
        os.mkdir('attachments')

    userName = raw_input('Enter your GMail username:')
    passwd = getpass.getpass('Enter your password: ')
    sender_email = 'khushi.ag27@gmail.com'
    try:
        imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
        typ, accountDetails = imapSession.login(userName, passwd)
        if typ != 'OK':
            print 'Not able to sign in!'
            raise
        
        imapSession.select('inbox')
        typ, data = imapSession.search(None,'UNSEEN', 'FROM', '"%s"' % sender_email)
        if typ != 'OK':
            print 'Error searching Inbox.'
            raise
        
        # Iterating over all emails
        for msgId in data[0].split():
            typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
            if typ != 'OK':
                print 'Error fetching mail.'
                raise

            emailBody = messageParts[0][1]
            mail = email.message_from_string(emailBody)
            
            for part in mail.walk():
                if part.get_content_maintype() == 'multipart':
                    # print part.as_string()
                    continue
                if part.get('Content-Disposition') is None:
                    # print part.as_string()
                    continue
                fileName = part.get_filename()
                print mail["Subject"]
                if mail["Subject"] == 'Fwd: Paytm Cash Summary for November 2017':
                    if bool(fileName):
                        filePath = os.path.join(detach_dir, 'attachments', fileName)
                        if not os.path.isfile(filePath) :
                            print fileName
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
        imapSession.close()
        imapSession.logout()
    except :
        print 'Not able to download all attachments.'

login_download_attachment()
excel_sheet()
