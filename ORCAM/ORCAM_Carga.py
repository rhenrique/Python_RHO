# encoding: utf-8
# encoding: iso-8859-1
# encoding: win-1252
import xlrd
import pyodbc
import time
import os
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import time
from ctypes import *
import smtplib
import mimetypes

emailfrom = "ORCAM@maxionwheels.com"
emailto = "rafael.oliveira@maxionwheels.com"
msg = MIMEMultipart()
msg["From"] = emailfrom
msg["To"] = emailto
msg["Subject"] = "ORCAM: Relatorio de Carga"

#----------------------------------------------------------------------
"""
Open and read an Excel file
"""
path = "Pasta1.xlsx"
#cdate = time.strftime("%Y%m")
cdate = '201711'
#mes = time.strftime("%m")
mes = '11'
cmonth = int(mes)
server = 'LIM-SQL12P1'
username = 'dsserver'
password = 'Maxion123@'
driver = '{SQL Server}'
port = '1433'
database = 'PROD_ORCAM_LMS'
param = 0
cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+password, autocommit=True)
cnn.setencoding(str, encoding='utf-8')
cnn.setencoding(unicode, encoding='utf-8', ctype=pyodbc.SQL_CHAR)
cursor = cnn.cursor()
book = xlrd.open_workbook(path)
#print number of sheets
#print book.nsheets

#print sheet names
#print book.sheet_names()

# get the first worksheet
out = open("output_file.txt", "w")
for read in cursor.execute("select * from ccusto where ccusto >= '0100'").fetchall():
	start = int(33 + cmonth)
	initial = 19 + cmonth
	end = int(34 + cmonth)
	c_custo = str(read.Ccusto[-3:])
	print c_custo
	try:
		first_sheet = book.sheet_by_name(c_custo)
	except:
		continue
	cell = first_sheet.cell(44,1)
	cells = first_sheet.col_slice(colx=1,start_rowx=44,end_rowx=159)
	for cell in cells:
		grupo = '0'+str(cell.value[2:-4])
		for c_grupo in cursor.execute("select * from Grupos").fetchall():
			#print "Planilha: " + grupo
			#print "Banco de Dados: " + c_grupo.Grupo
			if grupo == c_grupo.Grupo:
				chk = cursor.execute("select * from orcam where anomes=? and Ccusto=? and Grupo=?",cdate, '0'+c_custo, grupo).fetchone()
				#print cdate + c_custo + grupo
				if chk == None:
					print "Nao ha grupo " + grupo + " no CC " + c_custo
				else:
					#print "chegou aqui"
					#print initial
					cell2 = first_sheet.cell(44,initial)
					cells2 = first_sheet.col_slice(colx=30,start_rowx=start,end_rowx=end)
					for cell2 in cells2:
						forecast = float("%0.2f" % (cell2.value))
						#forecast = str("%0.2f" % (cell2.value))
						print forecast
						try:
							print "Paramentro: 0 CC: " + c_custo + " Grupo: " + grupo + " Ano/Mes: " + cdate + " Forecast: " + str(forecast)
							cursor.execute("update dbo.orcam set Tx_Forecast=0, Vr_Forecast=? where Ccusto=? and Grupo=? and Anomes=?", forecast, '0'+c_custo, grupo, cdate)
							cursor.commit()
							print >>out,"CC: " + c_custo + " Conta: " + grupo + " Ano/Mes: " + cdate + " Forecast: " + str(forecast)
						except Exception:
							print "Checar caracter"
		start += 1
		end += 1
out.close()
		
body = u"""<html><head> <font face="Calibri" size="3" color = "black"></head> <body> RELATÓRIO DE CARGA - ORÇAM - FORECAST</body></html>"""
#msg.attach(MIMEText(body))
msg.attach(MIMEText(body, 'html', 'utf-8'))
filename = "output_file.txt"
f = file(filename)
attachment = MIMEText(f.read())
attachment.add_header('Content-Disposition', 'attachment', filename=filename)           
msg.attach(attachment)

try:
    server = smtplib.SMTP('mail.maxionwheels.com',25)
    server.starttls()
    server.sendmail(emailfrom, emailto, msg.as_string())
    server.quit()
    print "Successfully sent email!"
except Exception:
    print "Error: unable to send email"
    server.quit()
		

#----------------------------------------------------------------------
