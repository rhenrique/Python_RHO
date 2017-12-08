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

server = 'LIM-SQL12P1'
username = 'dsserver'
password = 'Maxion123@'
driver = '{SQL Server}' # Driver necessário para conecta ao banco de dados do MSSQL
port = '1433'

#----------------------------------------------------------------------
"""
Open and read an Excel file
"""
cdate = time.strftime("%Y%m")
mes = time.strftime("%m")
cmonth = int(mes)
param = 0

#Definição do ambiente e do arquivo EXCEL a ser utilizado para carga
def ambiente():
	amb = raw_input("Escolha um ambiente:\n - [1] - ACO;\n - [2] - Aluminio.\n")
	if amb == '1':
		database = 'PROD_ORCAM_LMS'
	elif amb == '2':
		database = 'PROD_ORCAM_LMA'
	else:
		print u"Você digitou um ambiente errado... \nVamos começar de novo...."
		time.sleep(2)
		os.system('cls')
		ambiente()
	cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+password)
	cnn.setencoding(str, encoding='utf-8')
	cnn.setencoding(unicode, encoding='utf-8', ctype=pyodbc.SQL_CHAR)
	cursor = cnn.cursor()
	path = raw_input("Digite o nome do arquivo Excel.\nExemplo: forecast_aco.xlsx\n")
	book = xlrd.open_workbook(path)
	carga(cursor, database, book)

#Toda lógica do programa. 
def carga(cursor, database, book):
	out = open("FRSCT_" + database +".txt", "w")
	# 'for' para listar todo CC utilizado no ORÇAM
	for read in cursor.execute("select * from ccusto where ccusto >= '0100'").fetchall():
		#Variáveis para definir o inicio da leitura do EXCEL
		start = int(44) # Inicio da linha (valor incremental)
		initial = 19 + cmonth # Começa na coluna T + mês atual
		end = int(45) # Fim da linha (valor incremental)
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
						cells2 = first_sheet.col_slice(colx=initial,start_rowx=start,end_rowx=end)
						for cell2 in cells2:
							forecast = float("%0.2f" % (cell2.value))
							#forecast = str("%0.2f" % (cell2.value))
							print forecast
							try:
								print "Parametro: 0 CC: " + c_custo + " Grupo: " + grupo + " Ano/Mes: " + cdate + " Forecast: " + str(forecast)
								cursor.execute("update dbo.orcam set Tx_Forecast=0, Vr_Forecast=? where Ccusto=? and Grupo=? and Anomes=?", forecast, '0'+c_custo, grupo, cdate)
								cursor.commit()
								print >>out,"CC: " + c_custo + " Conta: " + grupo + " Ano/Mes: " + cdate + " Forecast: " + str(forecast)
							except Exception:
								print "Checar caracter"
			start += 1
			end += 1
	out.close()
			
	body = u"""<html><head> <font face="Calibri" size="3" color = "black"></head> <body> RELATÓRIO DE CARGA - ORÇAM - FORECAST</body></html>"""
	msg.attach(MIMEText(body, 'html', 'utf-8'))
	filename = "FRSCT_" + database +".txt"
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

ambiente()		

#----------------------------------------------------------------------
