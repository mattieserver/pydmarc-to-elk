from configparser import ConfigParser
from elasticsearch import Elasticsearch
from opensearchpy import OpenSearch
import sys

from datetime import datetime
import gzip
import base64
from zipfile import ZipFile
import json
import os
import time
import xml.etree.ElementTree as ET
import msal
import requests
import base64
import logging
import email
from email.message import EmailMessage
from io import BytesIO

CONFIG = ConfigParser()
CONFIG.read("Settings/config.ini")

CLIENT_ID = CONFIG.get("email", "client_id")
CLIENT_SECRET = CONFIG.get("email", "secret")
MAILBOX_ID = CONFIG.get("email", "mailbox_id")
DELETE_PROCESSED = CONFIG.getboolean("email", "delete_processed")

ELK_HOST = CONFIG.get("elk", "host")
ELK_PORT = CONFIG.get("elk", "port")
ELK_MODE = CONFIG.get("elk", "mode")
ELK_AUTH = CONFIG.get("elk", "auth")
ELK_USER = CONFIG.get("elk", "user")
ELK_PASSWORD = CONFIG.get("elk", "password")

authority = "https://login.microsoftonline.com/24133f96-2b38-4360-9546-979fc3caa9d7"
scopes = ["https://graph.microsoft.com/.default"]


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)


def generate_auth_string(user, token):
    authstr = f"user={user}\x01auth=Bearer {token}".encode("utf-8")
    return base64.b64encode(authstr)


class DMARCELK:
    __access_token = None

    def __init__(self, *args, **kwargs):
        requests.packages.urllib3.disable_warnings()
        self.__access_token = self.__setup_auth()
        self.__es = self.__setup_elk()

    def __setup_auth(self):
        app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)

        result = None
        result = app.acquire_token_silent(scopes, account=None)

        if not result:
            logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
            result = app.acquire_token_for_client(scopes=scopes)

        if "access_token" in result:
            logging.info("Got access Token")
        else:
            print(result.get("error"))
            print(result.get("error_description"))
            print(result.get("correlation_id"))

        return result["access_token"]

    def __setup_elk(self):
        es = None
        try:
            if ELK_AUTH == "yes":
                es_host = "https://{}:{}".format(ELK_HOST, ELK_PORT)
                es = OpenSearch(
                    [es_host],
                    http_auth=(ELK_USER, ELK_PASSWORD),
                    verify_certs=False,
                )
            else:
                es = OpenSearch([{"host": ELK_HOST, "port": ELK_PORT}])
            es.info()
        except Exception as ex:
            print(ex)
        return es

    def __handle_xml(self, xml_data, name, save_data=True):
        try:
            tree = ET.ElementTree(ET.fromstring(xml_data))
        except ET.ParseError:
            logging.error("Could not parse xml")
            return None
        if save_data:
            path_name = "data/processed/" + name + ".xml"
            path_name = path_name.replace("!", "#")
            tree.write(path_name)
            logging.info("Saved file to: {}".format(path_name))
        root = tree.getroot()
        output = {}
        output_rows = []
        if root.tag == "feedback":
            # version
            version = root.find("version")
            if version is not None:
                output["version"] = version.text
            else:
                output["version"] = ""

            # report_metadata
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

            # policy_published
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

            # record
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
            return output_rows

    def __upload_xml(self, xml_data):
        index_name = "dmarc-index-%s" % (time.strftime("%Y-%m"))
        if ELK_MODE == "read":
            self.__upload_xml_to_console(xml_data)
        elif ELK_MODE == "write":
            self.__upload_xml_to_elk(xml_data, index_name)
        else:
            print("ELK_MODE is not valid is: %s" % (ELK_MODE))

    def __upload_xml_to_elk(self, xml_data, index_name):
        logging.info("Writing data to index for {} records".format(len(xml_data)))
        for document in xml_data:
            self.__es.index(index=index_name, body=document)

    def __upload_xml_to_console(self, xml_data):
        logging.info("Writing data to log for {} records".format(len(xml_data)))
        for document in xml_data:
            json_data = json.dumps(document)
            logging.info(json_data)

    def __delete_mail_message(self, email_id):
        if DELETE_PROCESSED:
            delete_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_ID}/messages/{email_id}"
            delete_request = requests.delete(delete_url, headers={"Authorization": "Bearer " + self.__access_token})
            if delete_request.ok:
                logging.info("Deleted email: {}".format(email_id))
                return True
            else:
                logging.warning(
                    "Unable to delete email: {} with status code: {}".format(email_id, delete_request.status_code)
                )
                return False
        else:
            logging.info("Not deleting mail because of settings")
            return True

    def __read_gzip(self, att):
        try:
            file_content = gzip.decompress(att)
            return file_content
        except:
            logging.error("Unable to decompress gzip")
            return False

    def __read_zip(self, att):
        full_zip = ZipFile(BytesIO(att))

        if len(full_zip.namelist()) == 1:
            for zip_att in full_zip.namelist():
                zip_content = full_zip.open(zip_att).read()
                return zip_content
        else:
            logging.warning("Zip file contains to many zip files")

    def _handle_basic_attachements_data(self, contentType, attachment_name, data_bytes, email):
        xml_data = None
        if contentType == "application/gzip":
            logging.info("Got gzip attachment")
            clean_name = attachment_name.replace(".gz", "")
            clean_name = attachment_name.replace(".xml", "")
            unzipped = self.__read_gzip(data_bytes)
            if unzipped is not False:
                xml_data = self.__handle_xml(unzipped, clean_name)
                if xml_data is not None:
                    self.__upload_xml(xml_data)
            delete_message = True
            status = "ok"
        elif contentType == "application/zip":
            logging.info("Got zip attachment")
            clean_name = attachment_name.replace(".zip", "")
            unzipped = self.__read_zip(data_bytes)
            xml_data = self.__handle_xml(unzipped, clean_name)
            if xml_data is not None:
                self.__upload_xml(xml_data)
            delete_message = True
            status = "ok"
        elif contentType == "text/plain":
            logging.info("Got text attachment")
            logging.warning(
                "Got text attachment: {} ({})".format(email["subject"], email["sender"]["emailAddress"]["name"])
            )
            status = "not"
            delete_message = False
            # TODO: attachment is most likly an email, so we need to read the attachment form the content of the  attachments_value_request.content
        elif contentType == "application/octet-stream":
            logging.info("Got octet-stream attachment")
            logging.warning(
                "Got octet-stream attachment: {} ({})".format(
                    email["subject"], email["sender"]["emailAddress"]["name"]
                )
            )
            status = "not"
            delete_message = False
            # TODO: attachment is most likly an email, so we need to read the attachment form the content of the  attachments_value_request.content
        else:
            logging.warning("Got unkown content type:".format(contentType))
            status = "not"
            delete_message = False

        return status, delete_message

    def _handle_graph_extra_attachments(self, attachments_value_request, email):
        logging.info("Got onther email as email attachement")
        extra_attachment_json = attachments_value_request.json()
        att_status = "ok"
        delete_message = True
        if "item" in extra_attachment_json:
            if extra_attachment_json["item"]["@odata.type"] == "#microsoft.graph.message":
                if extra_attachment_json["item"]["hasAttachments"]:
                    for ms_extra_att in extra_attachment_json["item"]["attachments"]:
                        logging.info("looping over attachment")
                        self._handle_basic_attachements_data(
                            contentType=ms_extra_att["contentType"],
                            attachment_name=ms_extra_att["name"],
                            data_bytes=base64.b64decode(ms_extra_att["contentBytes"]),
                            email=email,
                        )
                        if ms_extra_att["contentType"] == "application/zip":
                            logging.info("Got zip")
        return att_status, delete_message

    def __run(self, url=None):
        endpoint = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_ID}/messages?$top=50&$select=sender,subject"
        if url is not None:
            endpoint = url
        r = requests.get(endpoint, headers={"Authorization": "Bearer " + self.__access_token})
        next_url = None
        if r.ok:
            logging.info("Retrieved emails successfully")
            data = r.json()
            if "@odata.nextLink" in data:
                next_url = data["@odata.nextLink"]
            for email in data["value"]:
                logging.debug(email["id"])
                if "sender" in email:
                    logging.info(
                        "Working on email with subject: {} and sender: {}".format(
                            email["subject"], email["sender"]["emailAddress"]["name"]
                        )
                    )
                else:
                    logging.warning("No sender in email")
                    logging.debug("No sender in email, dumping data: {}".format(email))
                    continue
                attachments_url = (
                    f'https://graph.microsoft.com/v1.0/users/{MAILBOX_ID}/messages/{email["id"]}/attachments'
                )
                attachments_request = requests.get(
                    attachments_url, headers={"Authorization": "Bearer " + self.__access_token}
                )
                if attachments_request.ok:
                    attachments_overiew_data = attachments_request.json()
                    delete_message = False
                    if len(attachments_overiew_data["value"]) > 1:
                        logging.warning("More than one attachment, mail might be delete without processing all data")
                    for attachment_overview in attachments_overiew_data["value"]:
                        logging.debug("attachment id ".format(attachment_overview["id"]))

                        extra_attachment_processing = False

                        if attachment_overview["@odata.type"] == "#microsoft.graph.itemAttachment":
                            extra_attachment_processing = True
                            attachments_value_url = f'https://graph.microsoft.com/v1.0/users/{MAILBOX_ID}/messages/{email["id"]}/attachments/{attachment_overview["id"]}/?$expand=microsoft.graph.itemattachment/item'
                        else:
                            attachments_value_url = f'https://graph.microsoft.com/v1.0/users/{MAILBOX_ID}/messages/{email["id"]}/attachments/{attachment_overview["id"]}/$value'

                        attachments_value_request = requests.get(
                            attachments_value_url, headers={"Authorization": "Bearer " + self.__access_token}
                        )
                        if attachments_value_request.ok:
                            if extra_attachment_processing:
                                att_status, delete_message = self._handle_graph_extra_attachments(
                                    attachments_value_request, email
                                )
                            else:
                                contentType = attachments_value_request.headers["content-type"]
                                if contentType == "application/octet-stream":
                                    if attachment_overview["name"].endswith(".gz"):
                                        logging.info("Override contentType for gzip")
                                        contentType = "application/gzip"
                                    elif attachment_overview["name"].endswith(".zip"):
                                        logging.info("Override contentType for zip")
                                        contentType = "application/zip"
                                    else:
                                        logging.warn("Override for contentType not supported")
                                att_status, delete_message = self._handle_basic_attachements_data(
                                    contentType, attachment_overview["name"], attachments_value_request.content, email
                                )

                    if delete_message:
                        self.__delete_mail_message(email["id"])
        else:
            logging.error("Could not retrieve mail")
        return next_url

    def start_run(self):
        next_url = self.__run()
        while next_url is not None:
            logging.info("Starting new Batch")
            next_url = self.__run(url=next_url)
        logging.info("Done mailbox")

    def reload_processed_folder(self):
        path = "data/processed/"
        files = os.listdir(path)
        for file_name in files:
            if file_name.endswith(".xml"):
                file_path = "data/processed/" + file_name
                file_object = open(file_path, "r").read()
                jsonreport = self.__handle_xml(file_object, file_name, False)
                self.__upload_xml(jsonreport)
