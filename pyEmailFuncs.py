#!/usr/bin/env python

from configparser import ConfigParser
import sys
import imaplib
import email
import email.header
import datetime
import gzip
import base64
import zipfile
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

    messages_data = data[0].split()
    messages_data_count = len(messages_data)
    print("Processing %i messages \n" % (messages_data_count))
    i = 0
    for num in messages_data:
        i += 1
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message %i" % (num))
            return

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        hdr_date = email.header.make_header(email.header.decode_header(msg['Date']))        
        print("Working on email: %i" % (i))          

        root_content = msg.get_content_type()
        if root_content == "multipart/mixed":
            for att in msg.get_payload(): 
                cnt_type = att.get_content_type()                
                if cnt_type == "text/plain":                
                    pass
                elif cnt_type == "multipart/alternative":
                    pass  
                elif cnt_type == "application/octet-stream":
                    handle_att("UNKNOWN", att)
                    pass           
                elif cnt_type == "application/zip":
                    handle_att("ZIP", att)
                elif cnt_type == "application/gzip":                
                    handle_att("GZIP", att)
                else:
                    print("unkown sub: %s" % (cnt_type))                    
        elif root_content == "application/zip":
            handle_att("ZIP", msg)
        else:
            print("unkown root: %s" % (cnt_type))     
        print("Done with email: %i \n" % (i))   

def handle_att(file_type, att):    
    att_name = att.get("Content-Description")    
    if att_name is not None:
        att_name = att_name.strip()
        att_name = att_name.replace(" ","")
        att_name_clean = att_name
        att_name = "data/" + file_type + "/" + att_name    
        open(att_name,'wb').write(att.get_payload(decode=True))
        handle_clean_att(file_type, att_name_clean, att_name)
    else:        
        att_dip = att.get("Content-Disposition")
        att_dip = att_dip.split(';')        
        for x in att_dip:
            x = x.replace("\r", "")
            x = x.replace("\n", "")
            x = x.replace("\t", "")
            x = x.replace(" ","")
            if x.startswith("filename"):                
                att_name = x.replace('"', '')
                att_name = att_name.replace("filename=","")
                att_name_clean = att_name
                if file_type == "UNKNOWN":
                    file_type_array = att_name_clean.split(".")
                    file_type_index = (len(file_type_array)) -1                    
                    if file_type_array[file_type_index] == "gz":
                        file_type = "GZIP"
                    else:
                        print("Unkown file type %s" % (file_type_array[file_type_index]))
                att_name = "data/" + file_type + "/" + att_name    
                open(att_name,'wb').write(att.get_payload(decode=True))
                handle_clean_att(file_type, att_name_clean, att_name)              

def handle_clean_att(file_type, att_name_clean, att_name):
    if file_type == "GZIP":
        f = gzip.open(att_name, 'rb')
        file_content = f.read()
        handle_xml(file_content, att_name_clean)        
        f.close()
    elif file_type == "ZIP":
        archive = zipfile.ZipFile(att_name, 'r')
        for a_file in archive.infolist():
            if a_file.filename.endswith(".xml"):
                xml = archive.read(a_file.filename)
                handle_xml(xml, att_name_clean)                
            else:
                print("File in ZIP is not a XML")       
    else:
        print("File Type not supported %s" % (file_type))
    print("--Filename: %s" % (att_name_clean))

def handle_xml(file, name):   
    path_name= "data/processed/" + name
    tree = ET.ElementTree(ET.fromstring(file))  
    #print(tree)
    tree.write(path_name)

M = imaplib.IMAP4_SSL(EMAIL_SERVER)

try:
    rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
except imaplib.IMAP4.error:
    print ("LOGIN FAILED!!! ")
    sys.exit(1)

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("Mailbox Login OK")
    process_mailbox(M)
    M.close()
else:
    #failed to get mailbox
    print("Cannot read mailbox")

M.logout()