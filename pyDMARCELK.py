from configparser import ConfigParser
from elasticsearch import Elasticsearch
import sys
import imaplib
import email
import email.header
from datetime import datetime
import gzip
import base64
import zipfile
import json
import os
import time
import xml.etree.ElementTree as ET

CONFIG = ConfigParser()
CONFIG.read('Settings/config.ini')

EMAIL_ACCOUNT = CONFIG.get('email', 'user')
EMAIL_FOLDER = CONFIG.get('email', 'reports_folder')
PROCESSED_FOLDER = CONFIG.get('email', 'processed_folder')
EMAIL_SERVER = CONFIG.get('email', 'host')
EMAIL_PASSWORD = CONFIG.get('email', 'password')
ELK_HOST = CONFIG.get('elk', 'host')
ELK_PORT = CONFIG.get('elk', 'port')
ELK_MODE = CONFIG.get('elk', 'mode')
es=Elasticsearch([{'host': ELK_HOST ,'port': ELK_PORT}])

class DMARCELK():
    __M = None
    __PROCESSED = []
    __messages_data = None
    __messages_data_start = None
    __data = None

    def __init__(self, *args, **kwargs):
        self.__setup_con()
        rv, self.__data = self.__M.search(None, "ALL")
        if rv != 'OK':
            print("No messages found!")

        self.__messages_data = self.__data[0].split()
        self.__messages_data_start = self.__messages_data

        messages_data_count = len(self.__messages_data)
        print("Processing %i messages \n" % (messages_data_count))

    def __setup_con(self):
        self.__M = imaplib.IMAP4_SSL(EMAIL_SERVER)
        self.__M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        self.__M.select(EMAIL_FOLDER, readonly = False)

    def start_run(self):
        try:
            self.__run()
            print("Done mailbox")
            self.__start_cleanup()
        except Exception as e:
            print(e)
            self.__M.logout()
            self.start_run()

    def __start_cleanup(self):
        try:
            self.__cleanup()
            print("Done Mailbox Cleanup")
        except Exception as e:
            print(e)
            self.__M.logout()
            self.__start_cleanup()

    def __cleanup(self):
        self.__setup_con()
        rv, new_data = self.__M.search(None, "ALL")
        if rv != 'OK':
            print("No messages found!")
        new_messages_data = new_data[0].split()
        for num in new_messages_data:
            uid = self.__getUID(self.__M,num)
            if uid in self.__PROCESSED:
                print("Message should be deteled")
                self.__M.store(num, '+FLAGS', '\\Deleted')
                self.__M.expunge()
        print("Messages cleaned up")
        items_left = self.__check_cleanup()
        if items_left != 0:
            self.__cleanup()
        else:
            print('Cleanup Check OK')

    def __check_cleanup(self):
        self.__M.logout()
        self.__setup_con()
        rv, new_data = self.__M.search(None, "ALL")
        if rv != 'OK':
            print("No messages found!")
        new_messages_data = new_data[0].split()
        new_messages_data_len = len(new_messages_data)
        count = 0
        if new_messages_data_len == 0:
            return 0
        else:
            for num in new_messages_data:
                uid = self.__getUID(self.__M,num)
                if uid in self.__PROCESSED:                    
                    count += 1                            
            return count

    def __getUID(self,M, num):
        resp, uid_data = M.fetch(num, "(UID)")
        if resp != 'OK':
            return
        uid_var = uid_data[0].decode()
        uid_var = uid_var.split("UID")
        uid_var = uid_var[1].replace(")","")
        uid_var = uid_var.replace(" ","")
        return uid_var

    def __process_mailbox(self, data, messages_data):
        for num in messages_data[:]:
            rv, data = self.__M.fetch(num, '(RFC822)')
            if rv != 'OK':
                return

            uid_var = self.__getUID(self.__M, num)

            if uid_var in self.__PROCESSED:
                continue

            msg = email.message_from_bytes(data[0][1])
            hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
            print("Working on email with uid %s and subject %s" % (uid_var,hdr))

            root_content = msg.get_content_type()
            if root_content == "multipart/mixed":
                for att in msg.get_payload():
                    cnt_type = att.get_content_type()
                    if cnt_type == "text/plain":
                        pass
                    elif cnt_type == "multipart/alternative":
                        pass
                    elif cnt_type == "application/octet-stream":
                        self.__handle_att("UNKNOWN", att)
                    elif cnt_type == "application/zip":
                        self.__handle_att("ZIP", att)
                    elif cnt_type == "application/gzip":
                        self.__handle_att("GZIP", att)
                    else:
                        print("unkown sub: %s" % (cnt_type))
            elif root_content == "application/zip":
                self.__handle_att("ZIP", msg)
            else:
                print("unkown root: %s" % (cnt_type))

            result_uid = self.__M.uid('COPY', uid_var, PROCESSED_FOLDER)            
            if result_uid[0] == "OK":
                pass
            else:
                print("Could not copy mail")
            self.__PROCESSED.append(uid_var)
            print("Done with email uid %s \n" % (uid_var))
            messages_data.remove(num)

    def __handle_att(self, file_type, att):
        att_name = att.get("Content-Description")
        if att_name is not None:
            att_name = att_name.strip()
            att_name = att_name.replace(" ","")
            self.__handle_clean_att(file_type, att_name, att)
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
                    self.__handle_clean_att(file_type, att_name, att)

    def __handle_clean_att(self, file_type, att_name_clean, att):
        att_name = att_name_clean
        if file_type != "UNKNOWN":
            pass
        else:
            file_type_array = att_name.split(".")
            file_type_index = (len(file_type_array)) -1
            file_extension = file_type_array[file_type_index]
            if file_extension == "gz":
                file_type = "GZIP"
            elif file_extension == "zip":
                file_type = "ZIP"
            else:
                print("Unkown file type %s" % (file_type_array[file_type_index]))
                return

        att_name = "data/" + file_type + "/" + att_name
        open(att_name,'wb').write(att.get_payload(decode=True))

        if file_type == "GZIP":
            f = gzip.open(att_name, 'rb')
            file_content = f.read()
            self.__handle_xml(file_content, att_name_clean, True)
            f.close()
        elif file_type == "ZIP":
            archive = zipfile.ZipFile(att_name, 'r')
            for a_file in archive.infolist():
                if a_file.filename.endswith(".xml"):
                    xml = archive.read(a_file.filename)
                    self.__handle_xml(xml, att_name_clean, True)
                else:
                    print("File in ZIP is not a XML")
        else:
            print("File Type not supported %s" % (file_type))

    def __handle_xml(self, file, name, new):
        tree = ET.ElementTree(ET.fromstring(file))
        if new:
            path_name= "data/processed/" + name + ".xml"        
            tree.write(path_name)
        root = tree.getroot()
        output = {}
        output_rows = []
        if root.tag == "feedback":
            #version
            version = root.find("version")
            if version is not None:
                output["version"] = version.text
            else:
                output["version"] = ""

            #report_metadata
            report_metadata = root.find("report_metadata")

            ##report_metadata - org_name
            org_name = report_metadata.find("org_name")
            if org_name is not None:
                output["report_metadata-org_name"] = org_name.text
            else:
                output["report_metadata-org_name"] = ""

            ##report_metadata - email
            email = report_metadata.find("email")
            if email is not None:
                output["report_metadata-email"] = email.text
            else:
                output["report_metadata-email"] = ""

            ##report_metadata - extra_contact_info
            extra_contact_info = report_metadata.find("extra_contact_info")
            if extra_contact_info is not None:
                output["report_metadata-extra_contact_info"] = extra_contact_info.text
            else:
                output["report_metadata-extra_contact_info"] = ""

            ##report_metadata - report_id
            report_id = report_metadata.find("report_id")
            if report_id is not None:
                output["report_metadata-report_id"] = report_id.text
            else:
                output["report_metadata-report_id"] = ""

            ##report_metadata - date_range
            date_range = report_metadata.find("date_range")
            if date_range is not None:
                begin = date_range.find("begin")
                end = date_range.find("end")

                ###report_metadata - date_range - begin
                if begin is not None:
                    output["report_metadata-date_range-begin"] = begin.text
                    int_begin = datetime.fromtimestamp(int(begin.text))
                    output["report_metadata-date_range-human_begin"] = int_begin.isoformat()
                else:
                    output["report_metadata-date_range-begin"] = ""
                    output["report_metadata-date_range-human_begin"] = ""

                ###report_metadata - date_range - end
                if end is not None:
                    output["report_metadata-date_range-end"] = end.text
                    int_end = datetime.fromtimestamp(int(end.text))
                    output["report_metadata-date_range-human_end"] = int_end.isoformat()
                else:
                    output["report_metadata-date_range-end"] = ""
                    output["report_metadata-date_range-human_end"] = ""
            else:
                output["report_metadata-date_range-end"] = ""
                output["report_metadata-date_range-begin"] = ""
                output["report_metadata-date_range-human_end"] = ""
                output["report_metadata-date_range-human_begin"] = ""

            #policy_published
            policy_published = root.find("policy_published")

            ##policy_published - domain
            domain = policy_published.find("domain")
            if domain is not None:
                output["policy_published-domain"] = domain.text
            else:
                output["policy_published-domain"] = ""

            ##policy_published - adkim
            adkim = policy_published.find("adkim")
            if adkim is not None:
                output["policy_published-adkim"] = adkim.text
            else:
                output["policy_published-adkim"] = ""

            ##policy_published - aspf
            aspf = policy_published.find("aspf")
            if aspf is not None:
                output["policy_published-aspf"] = aspf.text
            else:
                output["policy_published-aspf"] = ""

            ##policy_published - p
            p = policy_published.find("p")
            if p is not None:
                output["policy_published-p"] = p.text
            else:
                output["policy_published-p"] = ""

            ##policy_published - sp
            sp = policy_published.find("sp")
            if sp is not None:
                output["policy_published-sp"] = sp.text
            else:
                output["policy_published-sp"] = ""

            ##policy_published - pct
            pct = policy_published.find("pct")
            if pct is not None:
                output["policy_published-pct"] = pct.text
            else:
                output["policy_published-pct"] = ""

            #record            
            records = root.findall("record")
            for record in records:
                output_temp_record = output

                ##record - row
                row = record.find("row")
                if row is not None:
                    ###record - row - sourceip
                    source_ip = row.find("source_ip")
                    if source_ip is not None:
                        output_temp_record["row-source_ip"] = source_ip.text
                    else:
                        output_temp_record["row-source_ip"] = ""

                    ###record - row - count
                    count = row.find("count")
                    if count is not None:
                        output_temp_record["row-count"] = int(count.text)
                    else:
                        output_temp_record["row-count"] = ""
                    
                    ###record - row - policy_evaluated
                    policy_evaluated = row.find("policy_evaluated")

                    ####record - row - policy_evaluated - disposition
                    disposition = policy_evaluated.find("disposition")
                    if disposition is not None:
                        output_temp_record["row-policy_evaluated-disposition"] = disposition.text
                    else:
                        output_temp_record["row-policy_evaluated-disposition"] = ""

                    ####record - row - policy_evaluated - dkim
                    dkim = policy_evaluated.find("dkim")
                    if dkim is not None:
                        output_temp_record["row-policy_evaluated-dkim"] = dkim.text
                    else:
                        output_temp_record["row-policy_evaluated-dkim"] = ""

                    ####record - row - policy_evaluated - spf
                    spf = policy_evaluated.find("spf")
                    if spf is not None:
                        output_temp_record["row-policy_evaluated-spf"] = spf.text
                    else:
                        output_temp_record["row-policy_evaluated-spf"] = ""
                else:
                    output_temp_record["row-source_ip"] = ""
                    output_temp_record["row-count"] = ""
                    output_temp_record["row-policy_evaluated-disposition"] = ""
                    output_temp_record["row-policy_evaluated-dkim"] = ""
                    output_temp_record["row-policy_evaluated-spf"] = ""

                ##record - identifiers
                identifiers = record.find("identifiers")

                ###record - identifiers - header_from
                header_from = identifiers.find("header_from")
                if header_from is not None:
                    output_temp_record["identifiers-header_from"] = header_from.text
                else:
                    output_temp_record["identifiers-header_from"] = ""

                ###record - identifiers - envelope_from
                envelope_from = identifiers.find("envelope_from")
                if envelope_from is not None:
                    output_temp_record["identifiers-envelope_from"] = envelope_from.text
                else:
                    output_temp_record["identifiers-envelope_from"] = ""

                ##record - auth_results
                auth_results = record.find("auth_results")

                ###record - auth_results - spf
                spf = auth_results.find("spf")

                if spf is not None:
                    ###record - auth_results - spf - domain
                    spf_domain = spf.find("domain")
                    if spf_domain is not None:
                        output_temp_record["auth_results-spf-domain"] = spf_domain.text
                    else:
                        output_temp_record["auth_results-spf-domain"] = ""

                    ###record - auth_results - spf - result
                    spf_result = spf.find("result")
                    if spf_domain is not None:
                        output_temp_record["auth_results-spf-result"] = spf_result.text
                    else:
                        output_temp_record["auth_results-spf-result"] = ""

                    ###record - auth_results - spf - scope
                    spf_scope = spf.find("scope")
                    if spf_scope is not None:
                        output_temp_record["auth_results-spf-scope"] = spf_scope.text
                    else:
                        output_temp_record["auth_results-spf-scope"] = ""
                else:
                    output_temp_record["auth_results-spf-domain"] = ""
                    output_temp_record["auth_results-spf-result"] = ""
                    output_temp_record["auth_results-spf-scope"] = ""

                ###record - auth_results - dkim
                dkim = auth_results.find("dkim")

                if dkim is not None:
                    ###record - auth_results - dkim - domain
                    dkim_domain = dkim.find("domain")
                    if dkim_domain is not None:
                        output_temp_record["auth_results-dkim-domain"] = dkim_domain.text
                    else:
                        output_temp_record["auth_results-dkim-domain"] = ""

                    ###record - auth_results - dkim - result
                    dkim_result = dkim.find("result")
                    if dkim_result is not None:
                        output_temp_record["auth_results-dkim-result"] = dkim_result.text
                    else:
                        output_temp_record["auth_results-dkim-result"] = ""

                    ###record - auth_results - dkim - selector
                    dkim_selector = dkim.find("selector")
                    if dkim_selector is not None:
                        output_temp_record["auth_results-dkim-selector"] = dkim_selector.text
                    else:
                        output_temp_record["auth_results-dkim-selector"] = ""
                else:
                    output_temp_record["auth_results-dkim-domain"] = ""
                    output_temp_record["auth_results-dkim-result"] = ""
                    output_temp_record["auth_results-dkim-selector"] = ""
                
                if int_begin is not None:
                    output_temp_record["@timestamp"] = int_begin.isoformat()
                else:
                    output_temp_record["@timestamp"] = datetime.utcnow().isoformat()
                
                output_rows.append(output_temp_record)
                if ELK_MODE == "read":
                    json_data = json.dumps(output_temp_record)
                    print(json_data)                
                elif ELK_MODE == "write":
                    index_name = "dmarc-index-%s" % (time.strftime("%d-%m-%Y"))
                    es.index(index=index_name,doc_type='dmarc_report',body=output_temp_record)            
                else:
                    print("ELK_MODE is not valid is: %s" % (ELK_MODE))
                
        else:
            print("Unkown root tag for xml: %s" % (root.tag))

    def __run(self):
        self.__setup_con()
        self.__process_mailbox(self.__data, self.__messages_data)

    def reload_processed_folder(self):
        path = "data/processed/"
        files = os.listdir(path)
        for file_name in files:
            if file_name.endswith(".xml"):
                file_path = "data/processed/" + file_name
                file_object = open(file_path,"r").read()
                self.__handle_xml(file_object,file_name,False)
                