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
import json
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
					file_extension = file_type_array[file_type_index]					
					if file_extension == "gz":
						file_type = "GZIP"
					elif file_extension == "zip":
						file_type = "ZIP"
					else:
						print("Unkown file type %s" % (file_type_array[file_type_index]))
				if file_type != "UNKNOWN":
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

def handle_xml(file, name):
	path_name= "data/processed/" + name + ".xml"
	tree = ET.ElementTree(ET.fromstring(file))
	tree.write(path_name)
	root = tree.getroot()	
	output = {}
	if root.tag == "feedback":
		#version
		version = root.find("version")
		if version is not None:
			output["version"] = version.text
		else:
			output["version"] = ""

		#report_metadata
		report_metadata = root.find("report_metadata")
		output_report_metadata = {}

		##report_metadata - org_name
		org_name = report_metadata.find("org_name")
		if org_name is not None:
			output_report_metadata["org_name"] = org_name.text
		else:
			output_report_metadata["org_name"] = ""

		##report_metadata - email
		email = report_metadata.find("email")
		if email is not None:
			output_report_metadata["email"] = email.text
		else:
			output_report_metadata["email"] = ""

		##report_metadata - extra_contact_info
		extra_contact_info = report_metadata.find("extra_contact_info")
		if extra_contact_info is not None:
			output_report_metadata["extra_contact_info"] = extra_contact_info.text
		else:
			output_report_metadata["extra_contact_info"] = ""

		##report_metadata - report_id
		report_id = report_metadata.find("report_id")
		if report_id is not None:
			output_report_metadata["report_id"] = report_id.text
		else:
			output_report_metadata["report_id"] = ""

		##report_metadata - date_range
		date_range = report_metadata.find("date_range")
		output_date_range = {}
		if date_range is not None:
			begin = date_range.find("begin")
			end = date_range.find("end")			

			###report_metadata - date_range - begin
			if begin is not None:
				output_date_range["begin"] = begin.text
				int_begin = datetime.datetime.fromtimestamp(int(begin.text))			
				output_report_metadata["human_begin"] = int_begin.isoformat(' ')	
			else:
				output_date_range["begin"] = ""
				output_report_metadata["human_begin"] = ""

			###report_metadata - date_range - end
			if end is not None:
				output_date_range["end"] = end.text
				int_end = datetime.datetime.fromtimestamp(int(end.text))
				output_report_metadata["human_end"] = int_end.isoformat(' ')	
			else:
				output_date_range["end"] = ""
				output_report_metadata["human_end"] = ""
		else:
			output_date_range["end"] = ""
			output_date_range["begin"] = ""
			output_report_metadata["human_end"] = ""
			output_report_metadata["human_begin"] = ""
		output_report_metadata["date_range"] = output_date_range
		
		output["report_metadata"] = output_report_metadata

		#policy_published
		policy_published = root.find("policy_published")
		output_policy_published = {}

		##policy_published - domain
		domain = policy_published.find("domain")
		if domain is not None:
			output_policy_published["domain"] = domain.text
		else:
			output_policy_published["domain"] = ""

		##policy_published - adkim
		adkim = policy_published.find("adkim")
		if adkim is not None:
			output_policy_published["adkim"] = adkim.text
		else:
			output_policy_published["adkim"] = ""

		##policy_published - aspf
		aspf = policy_published.find("aspf")
		if aspf is not None:
			output_policy_published["aspf"] = aspf.text
		else:
			output_policy_published["aspf"] = ""

		##policy_published - p
		p = policy_published.find("p")
		if p is not None:
			output_policy_published["p"] = p.text
		else:
			output_policy_published["p"] = ""

		##policy_published - sp
		sp = policy_published.find("sp")
		if sp is not None:
			output_policy_published["sp"] = sp.text
		else:
			output_policy_published["sp"] = ""

		##policy_published - pct
		pct = policy_published.find("pct")
		if pct is not None:
			output_policy_published["pct"] = pct.text
		else:
			output_policy_published["pct"] = ""

		output["policy_published"] = output_policy_published

		#record
		records = root.findall("record")
		output_records = []
		for record in records:
			output_record = {}

			##record - row
			row = record.find("row")
			output_row = {}
			if row is not None:
				###record - row - sourceip
				source_ip = row.find("source_ip")
				if source_ip is not None:
					output_row["source_ip"] = source_ip.text
				else:
					output_row["source_ip"] = ""

				###record - row - count
				count = row.find("count")
				if count is not None:
					output_row["count"] = count.text
				else:
					output_row["count"] = ""

				###record - row - policy_evaluated
				policy_evaluated = row.find("policy_evaluated")
				output_policy_evaluated = {}

				####record - row - policy_evaluated - disposition
				disposition = policy_evaluated.find("disposition")
				if disposition is not None:
					output_policy_evaluated["disposition"] = disposition.text
				else:
					output_policy_evaluated["disposition"] = ""

				####record - row - policy_evaluated - dkim
				dkim = policy_evaluated.find("dkim")
				if dkim is not None:
					output_policy_evaluated["dkim"] = dkim.text
				else:
					output_policy_evaluated["dkim"] = ""

				####record - row - policy_evaluated - spf
				spf = policy_evaluated.find("spf")
				if spf is not None:
					output_policy_evaluated["spf"] = spf.text
				else:
					output_policy_evaluated["spf"] = ""

				output_row["policy_evaluated"] = output_policy_evaluated				
			else:
				pass	

			output_record["row"] = output_row

			##record - identifiers
			identifiers = record.find("identifiers")
			output_identifiers = {}

			###record - identifiers - header_from
			header_from = identifiers.find("header_from")
			if header_from is not None:
				output_identifiers["header_from"] = header_from.text
			else:
				output_identifiers["header_from"] = ""			

			###record - identifiers - envelope_from
			envelope_from = identifiers.find("envelope_from")
			if envelope_from is not None:
				output_identifiers["envelope_from"] = envelope_from.text
			else:
				output_identifiers["envelope_from"] = ""

			output_record["identifiers"] = output_identifiers

			##record - auth_results
			auth_results = record.find("auth_results")
			output_auth_results = {}

			###record - auth_results - spf
			spf = auth_results.find("spf")
			output_spf = {}
			if spf is not None:
				###record - auth_results - spf - domain
				spf_domain = spf.find("domain")
				if spf_domain is not None:
					output_spf["domain"] = spf_domain.text
				else:
					output_spf["domain"] = ""

				###record - auth_results - spf - result
				spf_result = spf.find("result")
				if spf_domain is not None:
					output_spf["result"] = spf_result.text
				else:
					output_spf["result"] = ""

				###record - auth_results - spf - scope
				spf_scope = spf.find("scope")
				if spf_scope is not None:
					output_spf["scope"] = spf_scope.text
				else:
					output_spf["scope"] = ""
			else:
				pass
			output_auth_results["spf"] = output_spf

			###record - auth_results - dkim
			dkim = auth_results.find("dkim")
			output_dkim = {}
			if dkim is not None:
				###record - auth_results - dkim - domain
				dkim_domain = dkim.find("domain")
				if dkim_domain is not None:
					output_dkim["domain"] = dkim_domain.text
				else:
					output_dkim["domain"] = ""

				###record - auth_results - dkim - result
				dkim_result = dkim.find("result")
				if dkim_result is not None:
					output_dkim["result"] = dkim_result.text
				else:
					output_dkim["result"] = ""

				###record - auth_results - dkim - selector
				dkim_selector = dkim.find("selector")
				if dkim_selector is not None:
					output_dkim["selector"] = dkim_selector.text
				else:
					output_dkim["selector"] = ""
			else:
				pass
			output_auth_results["dkim"] = output_dkim

			output_record["auth_results"] = output_auth_results

			output_records.append(output_record)

		output["records"] = output_records
	else:
		print("Unkown root tag for xml: %s" % (root.tag))
	json_data = json.dumps(output)
	print(json_data)

M = imaplib.IMAP4_SSL(EMAIL_SERVER)

try:
	rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
except imaplib.IMAP4.error:
	print ("LOGIN FAILED")
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