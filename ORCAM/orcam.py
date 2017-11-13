# coding=utf-8
import pyodbc
import sys
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import shutil
import subprocess
import time

cdate = time.strftime("%Y%m")
emailfrom = "rafael.oliveira@maxionwheels.com"
emailto = "rafael.oliveira@maxionwheels.com"
username = "oliveirh@maxionwheels.com"
password = "312414ra@#"
###################################
###Conexão com o Banco de Dados####
###################################
server = 'LIM-SQL12QA1'
database = 'TEST_ORCAM_LMS'
username = 'dsserver'
password = 'MaxionHOM123@'
driver = '{SQL Server}' # Driver you need to connect to the database
port = '1433'
cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+
                 ';PWD='+password)
cursor = cnn.cursor()
cursor.execute("select * from orcam")
for row in cursor.execute("select * from orcam where Anomes=?", cdate):
	#for owner in cursor.execute("select * from email where Ccusto=? and grupo=?", row.Ccusto, row.Grupo):
		#emailto = owner.email
		#msg["To"] = emailto
	try:
		percFor = (row.Vr_Forecast / 100) * 90
	except TypeError:
		percFor = row.Vr_Forecast
	if row.Vr_real > percFor and row.Vr_Forecast > 0:
		print(row.Ccusto, row.Grupo, row.Anomes)
		try:
			body = """<html><body> <font face="Calibri" size="3" color = "black"> Alertamos que os gastos com o CC """ + row.Ccusto + """ Grupo """ + row.Grupo + """, sob sua responsabilidade, excedeu 90% do previsto no Forecast para o mes """ + cdate + """ conforme abaixo: <br> <b>- Previsto:</b> """ + str(row.Vr_Forecast) + """ <br> <b>- Real:</b> """ + str(row.Vr_real) +  """ </body></html>"""
			msg = MIMEText(body, 'html')
			msg["From"] = emailfrom
			msg["Subject"] = "ORÇAM: Projeção está quase no limite."
			msg["To"] = emailto
			server = smtplib.SMTP('mail.maxionwheels.com',25)
			server.starttls()
			server.sendmail(emailfrom, emailto.split(';'), msg.as_string())
			server.quit()
			print "Successfully sent email!"
			server.close()
		except Exception:
			print "Error: unable to send email"
			server.quit()
			server.close()
	
		
	
	
	
	

	
