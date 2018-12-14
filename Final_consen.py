from mailmerge import MailMerge
import xlrd
import sys
import os
import comtypes.client
from comtypes.client import CreateObject
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from tkinter import messagebox
from email import encoders
import Login
import LoadUi
from datetime import date
import time
import logging
import logging.handlers

handler = logging.handlers.WatchedFileHandler(os.environ.get("LOGFILE", "E:\TUTORIALS_AND_CODINGS\Python_Coding_practice\CONSEN\Project_Final\EXECUITABLE\yourapp.log"))
formatter = logging.Formatter(logging.BASIC_FORMAT)
handler.setFormatter(formatter)
root = logging.getLogger()
root.setLevel(os.environ.get("LOGLEVEL", "INFO"))
root.addHandler(handler)

wdFormatPDF = 17

current_row = 0
sheet_num = 0
fromaddr = Login.username[0]
servername = Login.server_name[0]
portnum = Login.port_num[0]
password = Login.password[0]
greetings = Login.emailsubject[0]

msg = MIMEMultipart()
 
msg['From'] = fromaddr
msg['Subject'] = greetings

sheet_num = int(LoadUi.data_list[4])-1
src = str(LoadUi.data_list[1])
book = xlrd.open_workbook(src)
if((book.nsheets>=sheet_num+1) == False):
	print("Unable to find shet number provided.")
	sys.exit()
work_sheet = book.sheet_by_index(sheet_num)
finalList = []
headers = []
num_rows = work_sheet.nrows


try:

	print("Preparing to Merge.")
	while current_row < num_rows:
		dictVal=dict()
		if(current_row == 0):
			for col in range(work_sheet.ncols):
				headers.append(work_sheet.cell_value(current_row,col))
		else:
			for col in range(work_sheet.ncols):
				if col == (work_sheet.ncols - 1):
					dictVal.update({'Date':'{:%d-%b-%Y}'.format(date.today())})
				dictVal.update({headers[col]:work_sheet.cell_value(current_row,col)})
		if(current_row!=0):
			finalList.append(dictVal)
		
		current_row += 1
	print(len(finalList))
	print("Merge Operaton started.")
	for i in range(len(finalList)):
		greetings = Login.emailsubject[0]
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		#msg['To'] = toaddr
		msg['Subject'] = greetings
		first_name = finalList[i]['FirstName']
		last_name = finalList[i]['LastName']
		toaddr = finalList[i]['Email']
		msg['To'] = toaddr
		##########################################################################
		#Update Word
		##########################################################################
		document = MailMerge(str(LoadUi.data_list[0]))
		document.merge_pages([finalList[i]])
		print(f"Saving output file.{finalList[i]['FirstName']}")
		doc_file = f"{LoadUi.data_list[2]}\{finalList[i]['FirstName']}.docx"
		document.write(doc_file)
		##########################################################################
		#Convert Pdf
		##########################################################################
		print(f"Converting Pdf {finalList[i]['FirstName']}")
		pdf_output_path = f"{LoadUi.data_list[3]}\{finalList[i]['FirstName']}.pdf"
		word = comtypes.client.CreateObject('Word.Application')
		doc = word.Documents.Open(doc_file)
		print(f"Saving Pdf {finalList[i]['FirstName']}")
		doc.SaveAs(pdf_output_path, FileFormat=wdFormatPDF)
		doc.Close()
		word.Quit()
		###########################################################################
		#Sending Email
		###########################################################################
		print(toaddr)
		print(f"Sending Email {toaddr}")
		body = f"""Hi {first_name} {last_name},
				   
		Please find the attachment


				           """
		msg.attach(MIMEText(body, 'plain'))
	 
		filename = f"{finalList[i]['FirstName']}.pdf"
		attachment = open(f"{pdf_output_path}", "rb")

		part = MIMEBase('application', 'octet-stream')
		part.set_payload((attachment).read())
		encoders.encode_base64(part)
		part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
		 
		msg.attach(part)
		 
		server = smtplib.SMTP(servername,portnum )
		server.starttls()
		server.login(fromaddr, password)
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
		time.sleep(5)
		del msg
	server.quit()
		#############################################################################
		#
		#

		#############################################################################
		#
		#
	messagebox.showinfo("Completed", "operation completed Succesfully.")
	print("operation completed Succesfully.")
except Exception as e:
	logging.exception("Exceprion occured", exc_info=True)