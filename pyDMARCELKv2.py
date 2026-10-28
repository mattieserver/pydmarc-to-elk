from configparser import ConfigParser

import base64
import gzip
import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timezone
from io import BytesIO
from xml.etree.ElementTree import ElementTree, ParseError
from zipfile import BadZipFile, ZipFile

import msal
import requests
from defusedxml.ElementTree import fromstring as defused_fromstring
from opensearchpy import OpenSearch, helpers
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)

CONFIG_PATH = "Settings/config.ini"

REQUIRED_OPTIONS = {
    "elk": ["host", "port", "mode", "auth", "user", "password", "verify_certs"],
}

# Only the mailbox run needs these; reload_processed_folder works with [elk] alone.
MAILBOX_OPTIONS = ["tenant_id", "client_id", "secret", "mailbox_id", "delete_processed"]


def load_config():
    config = ConfigParser()
    if not config.read(CONFIG_PATH):
        raise SystemExit("Missing {}. Create it with: python3 writedefaultconf.py".format(CONFIG_PATH))

    missing = []
    for section, options in REQUIRED_OPTIONS.items():
        if not config.has_section(section):
            missing.append("[{}]".format(section))
            continue
        missing += ["[{}] {}".format(section, o) for o in options if not config.has_option(section, o)]
    if missing:
        raise SystemExit(
            "Incomplete {}, missing: {}. Run: python3 writedefaultconf.py".format(CONFIG_PATH, ", ".join(missing))
        )

    if config.get("elk", "mode") not in ("read", "write"):
        raise SystemExit("[elk] mode must be 'read' or 'write', got: {}".format(config.get("elk", "mode")))
    if config.get("elk", "auth") not in ("yes", "no"):
        raise SystemExit("[elk] auth must be 'yes' or 'no', got: {}".format(config.get("elk", "auth")))
    return config


def require_mailbox_config():
    """Checked only when we are about to talk to Graph, so a reload needs no mailbox credentials."""
    missing = [o for o in MAILBOX_OPTIONS if not CONFIG.has_option("email", o)]
    if missing:
        raise SystemExit(
            "Reading the mailbox needs [email] {} in {}. Run: python3 writedefaultconf.py".format(
                ", ".join(missing), CONFIG_PATH
            )
        )


CONFIG = load_config()

TENANT_ID = CONFIG.get("email", "tenant_id", fallback="")
CLIENT_ID = CONFIG.get("email", "client_id", fallback="")
CLIENT_SECRET = CONFIG.get("email", "secret", fallback="")
MAILBOX_ID = CONFIG.get("email", "mailbox_id", fallback="")
DELETE_PROCESSED = CONFIG.getboolean("email", "delete_processed", fallback=False)

ELK_HOST = CONFIG.get("elk", "host")
ELK_PORT = CONFIG.getint("elk", "port")
ELK_MODE = CONFIG.get("elk", "mode")
ELK_AUTH = CONFIG.get("elk", "auth")
ELK_USER = CONFIG.get("elk", "user")
ELK_PASSWORD = CONFIG.get("elk", "password")
ELK_VERIFY_CERTS = CONFIG.getboolean("elk", "verify_certs")

AUTHORITY = "https://login.microsoftonline.com/{}".format(TENANT_ID)
SCOPES = ["https://graph.microsoft.com/.default"]
GRAPH_MAILBOX = "https://graph.microsoft.com/v1.0/users/{}".format(MAILBOX_ID)
GRAPH_BATCH = "https://graph.microsoft.com/v1.0/$batch"
GRAPH_BATCH_SIZE = 20

# (method, result tag, document field) for auth_results, precomputed to keep it out of the record loop.
AUTH_RESULT_FIELDS = [
    (method, tag, "auth_results-{}-{}".format(method, tag))
    for method, tags in (("spf", ("domain", "result", "scope")), ("dkim", ("domain", "result", "selector")))
    for tag in tags
]

PROCESSED_DIR = "data/processed"
UNSAFE_NAME = re.compile(r"[^A-Za-z0-9._-]")
MAX_DECOMPRESSED_BYTES = 64 * 1024 * 1024
# A failed bulk chunk reports an error per document; only the first few are worth logging.
LOGGED_BULK_ERRORS = 5


def text_of(parent, tag, default=""):
    """Text of <tag> under parent, or default when either is missing."""
    if parent is None:
        return default
    child = parent.find(tag)
    if child is None or child.text is None:
        return default
    return child.text


def int_of(parent, tag, default=""):
    raw = text_of(parent, tag)
    if raw == "":
        return default
    try:
        return int(raw)
    except ValueError:
        logging.warning("Expected a number in <{}>, got: {}".format(tag, raw))
        return default


def iso_utc(epoch):
    if epoch == "":
        return ""
    return datetime.fromtimestamp(epoch, tz=timezone.utc).isoformat()


def strip_suffixes(name, suffixes):
    for suffix in suffixes:
        if name.lower().endswith(suffix):
            name = name[: -len(suffix)]
    return name


def processed_path(name):
    """Path under PROCESSED_DIR for an attacker-controlled attachment name."""
    clean = UNSAFE_NAME.sub("#", os.path.basename(name)).strip(".") or "report"
    path = os.path.join(PROCESSED_DIR, clean + ".xml")
    root = os.path.realpath(PROCESSED_DIR)
    if os.path.commonpath([os.path.realpath(path), root]) != root:
        raise ValueError("Refusing to write outside {}: {}".format(PROCESSED_DIR, name))
    return path


def strip_namespace(root):
    """RFC 9990 (dmarc-2.0) reports are namespace-qualified, RFC 7489 ones are not."""
    if not root.tag.startswith("{"):
        return
    for element in root.iter():
        if element.tag.startswith("{"):
            element.tag = element.tag.split("}", 1)[1]


def read_limited(fileobj, limit):
    data = fileobj.read(limit + 1)
    if len(data) > limit:
        raise ValueError("Decompressed payload exceeds {} bytes".format(limit))
    return data


class DMARCELK:
    def __init__(self):
        self.__app = None
        self.__session = None
        self.__es = self.__setup_elk()

    def __setup_graph(self):
        """Graph is set up on demand: reload_processed_folder never reads the mailbox."""
        require_mailbox_config()
        self.__app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )
        self.__session = self.__setup_session()
        self.__token()

    def __setup_session(self):
        session = requests.Session()
        retry = Retry(
            total=5,
            backoff_factor=1,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET", "DELETE", "POST"]),
            respect_retry_after_header=True,
        )
        session.mount("https://", HTTPAdapter(max_retries=retry))
        return session

    def __token(self):
        result = self.__app.acquire_token_silent(SCOPES, account=None)
        if not result:
            logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
            result = self.__app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" not in result:
            logging.error(
                "Could not acquire a token: {} - {} ({})".format(
                    result.get("error"), result.get("error_description"), result.get("correlation_id")
                )
            )
            sys.exit(1)
        return result["access_token"]

    def __headers(self):
        return {"Authorization": "Bearer " + self.__token()}

    def __setup_elk(self):
        if ELK_MODE != "write":
            logging.info("ELK_MODE is {}, not connecting to OpenSearch".format(ELK_MODE))
            return None
        if ELK_AUTH == "yes":
            es_host = "https://{}:{}".format(ELK_HOST, ELK_PORT)
            es = OpenSearch(hosts=[es_host], http_auth=(ELK_USER, ELK_PASSWORD), verify_certs=ELK_VERIFY_CERTS)
        else:
            es = OpenSearch(hosts=[{"host": ELK_HOST, "port": ELK_PORT}])
        es.info()
        logging.info("Connected to OpenSearch")
        return es

    def __handle_xml(self, xml_data, name, save_data=True):
        """Parse one aggregate report. Returns a list of documents, or None on failure."""
        if xml_data is None:
            logging.error("No xml data to parse for: {}".format(name))
            return None
        try:
            root = defused_fromstring(xml_data)
        except (ParseError, ValueError, TypeError) as ex:
            logging.error("Could not parse xml for {}: {}".format(name, ex))
            return None

        if save_data:
            try:
                path_name = processed_path(name)
            except ValueError as ex:
                logging.error(str(ex))
                return None
            ElementTree(root).write(path_name)
            logging.info("Saved file to: {}".format(path_name))

        strip_namespace(root)
        if root.tag != "feedback":
            logging.error("Expected a <feedback> report for {}, got: <{}>".format(name, root.tag))
            return None

        output = {"version": text_of(root, "version")}

        # report_metadata
        report_metadata = root.find("report_metadata")
        for tag in ("org_name", "email", "extra_contact_info", "report_id"):
            output["report_metadata-" + tag] = text_of(report_metadata, tag)

        ##report_metadata - date_range
        date_range = report_metadata.find("date_range") if report_metadata is not None else None
        for tag in ("begin", "end"):
            output["report_metadata-date_range-" + tag] = text_of(date_range, tag)
            output["report_metadata-date_range-human_" + tag] = iso_utc(int_of(date_range, tag))

        # policy_published
        policy_published = root.find("policy_published")
        for tag in ("domain", "adkim", "aspf", "p", "sp", "pct", "discovery_method", "testing"):
            output["policy_published-" + tag] = text_of(policy_published, tag)

        report_timestamp = output["report_metadata-date_range-human_begin"] or datetime.now(timezone.utc).isoformat()

        # record
        output_rows = []
        for record in root.findall("record"):
            row_record = dict(output)

            ##record - row
            row = record.find("row")
            row_record["row-source_ip"] = text_of(row, "source_ip")
            row_record["row-count"] = int_of(row, "count")

            ###record - row - policy_evaluated
            policy_evaluated = row.find("policy_evaluated") if row is not None else None
            for tag in ("disposition", "dkim", "spf"):
                row_record["row-policy_evaluated-" + tag] = text_of(policy_evaluated, tag)

            ##record - identifiers
            identifiers = record.find("identifiers")
            for tag in ("header_from", "envelope_from"):
                row_record["identifiers-" + tag] = text_of(identifiers, tag)

            ##record - auth_results (a record may carry several spf/dkim results)
            auth_results = record.find("auth_results")
            found = {}
            for method, tag, field in AUTH_RESULT_FIELDS:
                if method not in found:
                    found[method] = auth_results.findall(method) if auth_results is not None else []
                row_record[field] = [text_of(r, tag) for r in found[method]]

            row_record["@timestamp"] = report_timestamp
            output_rows.append(row_record)
        return output_rows

    def __document_id(self, document, position):
        return "{}-{}-{}".format(
            document["report_metadata-org_name"], document["report_metadata-report_id"], position
        )

    def __upload_xml(self, xml_data):
        if not xml_data:
            logging.info("Report contains no records, nothing to upload")
            return True
        if ELK_MODE == "read":
            return self.__upload_xml_to_console(xml_data)
        if ELK_MODE == "write":
            index_name = "dmarc-index-%s" % (time.strftime("%Y-%m"))
            return self.__upload_xml_to_elk(xml_data, index_name)
        logging.error("ELK_MODE is not valid is: %s" % (ELK_MODE))
        return False

    def __upload_xml_to_elk(self, xml_data, index_name):
        logging.info("Writing data to index for {} records".format(len(xml_data)))
        actions = [
            {"_index": index_name, "_id": self.__document_id(document, position), "_source": document}
            for position, document in enumerate(xml_data)
        ]
        try:
            indexed, errors = helpers.bulk(self.__es, actions, raise_on_error=False, raise_on_exception=False)
        except Exception as ex:
            logging.error("Could not write documents to {}: {}".format(index_name, ex))
            return False
        for error in errors[:LOGGED_BULK_ERRORS]:
            logging.error("Could not write document to {}: {}".format(index_name, error))
        if errors:
            logging.error("{} of {} documents rejected by {}".format(len(errors), len(actions), index_name))
            return False
        logging.info("Indexed {} documents into {}".format(indexed, index_name))
        return True

    def __upload_xml_to_console(self, xml_data):
        logging.info("Writing data to log for {} records".format(len(xml_data)))
        for document in xml_data:
            logging.info(json.dumps(document))
        return True

    def __delete_mail_messages(self, email_ids):
        if not DELETE_PROCESSED:
            logging.info("Not deleting {} mails because of settings".format(len(email_ids)))
            return
        for start in range(0, len(email_ids), GRAPH_BATCH_SIZE):
            chunk = email_ids[start : start + GRAPH_BATCH_SIZE]
            batch = {
                "requests": [
                    {"id": str(i), "method": "DELETE", "url": "/users/{}/messages/{}".format(MAILBOX_ID, email_id)}
                    for i, email_id in enumerate(chunk)
                ]
            }
            response = self.__session.post(GRAPH_BATCH, headers=self.__headers(), json=batch)
            if not response.ok:
                logging.error("Batch delete failed with status code: {}".format(response.status_code))
                continue
            for result in response.json().get("responses", []):
                email_id = chunk[int(result["id"])]
                if 200 <= result["status"] < 300:
                    logging.info("Deleted email: {}".format(email_id))
                else:
                    logging.warning(
                        "Unable to delete email: {} with status code: {}".format(email_id, result["status"])
                    )

    def __read_gzip(self, att):
        try:
            with gzip.GzipFile(fileobj=BytesIO(att)) as unzipped:
                return read_limited(unzipped, MAX_DECOMPRESSED_BYTES)
        except (OSError, EOFError, ValueError) as ex:
            logging.error("Unable to decompress gzip: {}".format(ex))
            return None

    def __read_zip(self, att):
        try:
            with ZipFile(BytesIO(att)) as full_zip:
                names = full_zip.namelist()
                if len(names) != 1:
                    logging.warning("Expected exactly one file in the zip, got {}".format(len(names)))
                    return None
                with full_zip.open(names[0]) as member:
                    return read_limited(member, MAX_DECOMPRESSED_BYTES)
        except (BadZipFile, OSError, ValueError) as ex:
            logging.error("Unable to read zip: {}".format(ex))
            return None

    def _handle_basic_attachements_data(self, contentType, attachment_name, data_bytes, message):
        """Returns True only when the report was parsed and uploaded, i.e. the mail may be deleted."""
        if contentType == "application/gzip":
            logging.info("Got gzip attachment")
            clean_name = strip_suffixes(attachment_name, (".gz", ".xml"))
            unzipped = self.__read_gzip(data_bytes)
        elif contentType == "application/zip":
            logging.info("Got zip attachment")
            clean_name = strip_suffixes(attachment_name, (".zip", ".xml"))
            unzipped = self.__read_zip(data_bytes)
        else:
            # TODO: text/plain and application/octet-stream are most likely an email, so we need to
            # read the attachment from the content of the attachments_value_request.content
            logging.warning(
                "Got unsupported content type {} on: {} ({})".format(
                    contentType, message["subject"], message["sender"]["emailAddress"]["name"]
                )
            )
            return False

        if unzipped is None:
            return False
        xml_data = self.__handle_xml(unzipped, clean_name)
        if xml_data is None:
            return False
        return self.__upload_xml(xml_data)

    def _handle_graph_extra_attachments(self, attachments_value_request, message):
        logging.info("Got another email as email attachment")
        item = attachments_value_request.json().get("item")
        if item is None or item.get("@odata.type") != "#microsoft.graph.message":
            logging.warning("Item attachment is not a message, leaving the mail in the mailbox")
            return False
        if not item.get("hasAttachments"):
            logging.warning("Attached message has no attachments, leaving the mail in the mailbox")
            return False
        nested_attachments = item.get("attachments")
        if not nested_attachments:
            logging.warning("Attached message reports attachments but none were expanded by Graph")
            return False

        delete_message = True
        for ms_extra_att in nested_attachments:
            logging.info("looping over attachment")
            handled = self._handle_basic_attachements_data(
                contentType=ms_extra_att["contentType"],
                attachment_name=ms_extra_att["name"],
                data_bytes=base64.b64decode(ms_extra_att["contentBytes"]),
                message=message,
            )
            delete_message = handled and delete_message
        return delete_message

    def __process_attachment(self, message, attachment_overview):
        attachment_url = "{}/messages/{}/attachments/{}".format(
            GRAPH_MAILBOX, message["id"], attachment_overview["id"]
        )
        is_item_attachment = attachment_overview["@odata.type"] == "#microsoft.graph.itemAttachment"
        if is_item_attachment:
            attachments_value_url = attachment_url + "/?$expand=microsoft.graph.itemattachment/item"
        else:
            attachments_value_url = attachment_url + "/$value"

        attachments_value_request = self.__session.get(attachments_value_url, headers=self.__headers())
        if not attachments_value_request.ok:
            logging.warning(
                "Unable to read attachment: {} with status code: {}".format(
                    attachment_overview["id"], attachments_value_request.status_code
                )
            )
            return False
        if is_item_attachment:
            return self._handle_graph_extra_attachments(attachments_value_request, message)

        contentType = attachments_value_request.headers.get("content-type", "").split(";")[0].strip().lower()
        if contentType == "application/octet-stream":
            if attachment_overview["name"].endswith(".gz"):
                logging.info("Override contentType for gzip")
                contentType = "application/gzip"
            elif attachment_overview["name"].endswith(".zip"):
                logging.info("Override contentType for zip")
                contentType = "application/zip"
            else:
                logging.warning("Override for contentType not supported")
        return self._handle_basic_attachements_data(
            contentType, attachment_overview["name"], attachments_value_request.content, message
        )

    def __process_message(self, message):
        """Returns True only when every attachment was processed, i.e. the mail may be deleted."""
        logging.debug(message["id"])
        if "sender" not in message:
            logging.warning("No sender in email")
            logging.debug("No sender in email, dumping data: {}".format(message))
            return False
        logging.info(
            "Working on email with subject: {} and sender: {}".format(
                message["subject"], message["sender"]["emailAddress"]["name"]
            )
        )

        attachments_url = "{}/messages/{}/attachments".format(GRAPH_MAILBOX, message["id"])
        attachments_request = self.__session.get(attachments_url, headers=self.__headers())
        if not attachments_request.ok:
            logging.warning(
                "Unable to list attachments for: {} with status code: {}".format(
                    message["id"], attachments_request.status_code
                )
            )
            return False

        attachments_overiew_data = attachments_request.json()["value"]
        if not attachments_overiew_data:
            logging.warning("Email has no attachments, leaving it in the mailbox")
            return False

        delete_message = True
        for attachment_overview in attachments_overiew_data:
            logging.debug("attachment id {}".format(attachment_overview["id"]))
            handled = self.__process_attachment(message, attachment_overview)
            delete_message = handled and delete_message
        return delete_message

    def __run(self, to_delete, url=None):
        endpoint = url
        if endpoint is None:
            endpoint = "{}/messages?$top=50&$select=sender,subject".format(GRAPH_MAILBOX)
        r = self.__session.get(endpoint, headers=self.__headers())
        if not r.ok:
            logging.error("Could not retrieve mail, status code: {}".format(r.status_code))
            return None
        logging.info("Retrieved emails successfully")
        data = r.json()
        for message in data["value"]:
            if self.__process_message(message):
                to_delete.append(message["id"])
        return data.get("@odata.nextLink")

    def start_run(self):
        self.__setup_graph()
        # Deleting while paginating shifts the result window, so collect first and delete afterwards.
        to_delete = []
        next_url = self.__run(to_delete)
        while next_url is not None:
            logging.info("Starting new Batch")
            next_url = self.__run(to_delete, url=next_url)
        logging.info("Done mailbox, {} emails processed".format(len(to_delete)))
        self.__delete_mail_messages(to_delete)

    def reload_processed_folder(self):
        for file_name in sorted(os.listdir(PROCESSED_DIR)):
            if not file_name.endswith(".xml"):
                continue
            file_path = os.path.join(PROCESSED_DIR, file_name)
            with open(file_path, "rb") as report_file:
                file_object = report_file.read()
            jsonreport = self.__handle_xml(file_object, file_name, False)
            if jsonreport is None:
                logging.warning("Skipping: {}".format(file_path))
                continue
            self.__upload_xml(jsonreport)
