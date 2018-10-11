#!/usr/bin/env python

from configparser import ConfigParser
import sys
import imaplib
import email
import email.header
import datetime
import gzip
import xml.etree.ElementTree as ET

CONFIG = ConfigParser()
CONFIG.read('Settings/config.ini')

EMAIL_ACCOUNT = CONFIG.get('email', 'user')
EMAIL_FOLDER = CONFIG.get('email', 'reports_folder')
EMAIL_SERVER = CONFIG.get('email', 'host')
EMAIL_PASSWORD = CONFIG.get('email', 'password')

def process_mailbox(M):
    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        subject = str(hdr)

        #print(msg.get_payload())        
        for att in msg.get_payload():
            cnt_type = att.get_content_type()            
            if cnt_type == "text/plain":                
                pass                
            elif cnt_type == "application/zip":
                handle_att("ZIP", att)
            elif cnt_type == "application/gzip":                
                handle_att("GZIP", att)
            else:
                print("unkown: %s" % (cnt_type))
            

def handle_att(file_type, att):
    att_name = att.get("Content-Description")
    att_name = att_name.strip()
    att_name = att_name.replace(" ","")
    att_name = "data/" + file_type + "/" + att_name    
    open(att_name,'wb').write(att.get_payload(decode=True))

    if file_type == "GZIP":
        f = gzip.open(att_name, 'rb')
        file_content = f.read()
        handle_xml(file_content)        
        f.close()

def handle_xml(file):    
    tree = ET.ElementTree(ET.fromstring(file))    
    print(tree.getroot())


M = imaplib.IMAP4_SSL(EMAIL_SERVER)

try:
    rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
except imaplib.IMAP4.error:
    print ("LOGIN FAILED!!! ")
    sys.exit(1)

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("Processing mailbox...\n")
    process_mailbox(M)
    M.close()
else:
    #failed to get mailbox
    print("Cannot read mailbox")

M.logout()