# encoding: utf-8
# encoding: iso-8859-1
# encoding: win-1252
import pyodbc
import sys
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
import os
from distutils.core import setup
from ConfigParser import SafeConfigParser
import glob

#Variável que recebe o ano e o mês atual no formato 'YYYYMM'
cdate = time.strftime("%Y%m")
ins_date = time.strftime("%Y%m%d")

##########################################################
### Configuração de conexão usado para conecta ao banco ##
### LIB: PYODBC							    			##
### Documentação: https://goo.gl/jL4nAz					##
### Instalação: https://goo.gl/7ohaCq
##														##
##########################################################

## Versão: 0.1 - Concepção de um novo Sistema integrado com banco de dados/Orçam
## Versão 0.2 - Alterado para considerar por CC;
## Versão 0.3 - Bug fixes: considerar apenas enviar email quando há registro na string body
## Versão 0.4 - Alterado corpo do e-mail, bug fixes
## Versão Main 1.0 - Compilado para uso e adicionado logs
## Versão 1.1 - Adicionado controle ConfigParser - arquivo de configuração
 
parser = SafeConfigParser()
parser.read('ORCAM_EMAIL.ini')
 
emailfrom = "ORCAM_report@maxionwheels.com"
server = parser.get('amb_orcam', 'server')
#database = 'TEST_ORCAM_LMS'
username = parser.get('amb_orcam', 'user_db')
password = 'MaxionHOM123@'
driver = '{SQL Server}' # Driver necessário para conecta ao banco de dados do MSSQL
port = '1433'
forecast = parser.get('forecast','forecast')
print forecast

		
def app(cnn, cursor):
	row = cursor.execute("select * from Grupos")
	i = 0
	body = []
	for row in cursor.fetchall():
		body = []
		for owner in cursor.execute("select * from orcam INNER JOIN Grupos ON orcam.Grupo = Grupos.Grupo INNER JOIN ccustob ON ccustob.CcustoB = orcam.Ccusto where orcam.Grupo=? and anomes=?", row.Grupo, cdate):
			if forecast == 1:
				fb_value = owner.Vr_Forecast
			elif forecast == 0:
				fb_value = owner.Vr_Budget
			try:
				percFor = (fb_value / 100) * 80
			except TypeError:
				percFor = 0
			if owner.Vr_real > percFor and fb_value > 0:
				perc = (owner.Vr_real / fb_value) * 100
				print "- Conta: " + owner.Grupo + " CC: " + owner.Ccusto + u" Ano/Mês: " + owner.Anomes
				sub = u"ORÇAM: Projetado x Real - Aço"
				body.append(u"""<html> <head> <font face="Calibri" size="3" color = "black"> </head><body> <b>- CC: </b>""" + owner.Ccusto[-3:] + u""" - """ + str(owner[19].encode('ascii', 'ignore')) + u"""		<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; - Previsto:</b> """ + str("%0.2f" % (fb_value)).replace('.',',') + """&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>- Real:</b> """ + str("%0.2f" % (owner.Vr_real)).replace('.',',') +  u"""<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;%</b>: """ + str(perc).split('.')[0] + u""" % </font></body></html>""")
				grupo = str(owner[16].encode('ascii', 'ignore'))
		if body:
			sent = cursor.execute("""DECLARE @AllValues VARCHAR(4000)
									SELECT @AllValues = COALESCE(@AllValues + ';', '') + email 
									from grupo_email inner join grupo_resp on grupo_resp.resp = grupo_email.id where grupo_resp.grupo=?
									select @AllValues as email_sent""", owner.Grupo).fetchval()
			emailto = str(sent)
			body2 = u"""<html><head> <font face="Calibri" size="3" color = "black"></head> <body> Alertamos que os gastos da Conta Contábil <b>""" + owner.Grupo[-3:] +  """ - """ + str(grupo) + u"""</b> sob sua responsabilidade, excedeu 80% do previsto no Forecast para o ano/mês <b>""" + cdate + """</b> conforme detalhado abaixo: <br><br></font></body></html>"""
			if sent == None:
				print "Nao ha responsavel para a Conta: " + owner.Grupo 
			else:
				ins_grupo = "Conta: " + owner.Ccusto
				emailfrom = "ORCAM_report@maxionwheels.com"
				msgfinal = ''.join(body)
				msglala = body2 + msgfinal
				msg = MIMEText(msglala, 'html', 'utf-8')
				msg["From"] = emailfrom
				msg["Subject"] = sub
				msg["To"] = emailto
				server = smtplib.SMTP('mail.maxionwheels.com',25)
				server.starttls()
				try:
					server.sendmail(emailfrom, emailto.split(';'), msg.as_string())
				except SocketError:
					print u"Nao foi possivel conectar ao servidor"
					cursor.execute("INSERT INTO logs VALUES (?, ?, ?, 'Erro ao conectar no servidor');", ins_date, ins_grupo, sent)
					cursor.commit()
					msglala = ""
				else:
					print u"Email enviado com sucesso para " + sent
					cursor.execute("INSERT INTO logs VALUES (?, ?, ?, 'Email enviado com sucesso');", ins_date, ins_grupo, sent)
					cursor.commit()
					msglala = ""
		else:
			print u"O real não foi excedido em nenhum CC para a Conta: " + owner.Grupo[-3:]
	os.system('cls')
	print u"Foi finalizado o envio de informações para Conta Contabil."
	print u"Será iniciado o envio de e-mail para CC."	
	#server.quit()
	time.sleep(5)
	ccresp(cnn, cursor)

def ccresp(cnn, cursor):
	row = cursor.execute("select * from ccusto where ccusto >= '0100'")
	i = 0
	body = []
	for row in cursor.fetchall():
		body = []
		for owner in cursor.execute("select * from orcam INNER JOIN Ccusto ON orcam.Ccusto = Ccusto.Ccusto INNER JOIN grupos ON grupos.Grupo = orcam.Grupo where orcam.Ccusto=? and anomes=?", row.Ccusto, cdate):
			if forecast == 1:
				fb_value = owner.Vr_Forecast
			elif forecast == 0:
				fb_value = owner.Vr_Budget
			try:
				percFor = (fb_value / 100) * 80
			except TypeError:
				percFor = 0
			if owner.Vr_real > percFor and fb_value > 0:
				perc = (owner.Vr_real / fb_value) * 100
				print "- CC: " + owner.Ccusto + " Grupo: " + owner.Grupo + u" Ano/Mês: " + owner.Anomes
				sub = u"ORÇAM: Projetado x Real - Aço"
				body.append("""<html> <head> <font face="Calibri" size="3" color = "black"> </head><body> <b>- Conta: </b>""" + owner.Grupo[-3:] + """ - """ + str(owner[17].encode('ascii', 'ignore')) +  u""" <b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; - Previsto:</b> """ + str("%0.2f" % (fb_value)).replace('.',',') + """&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>- Real:</b> """ + str("%0.2f" % (owner.Vr_real)).replace('.',',') +  u"""<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;%</b>: """ + str(perc).split('.')[0] + u""" % </font></body></html>""")
		if body:
			sent = cursor.execute("""DECLARE @AllValues VARCHAR(4000)
									SELECT @AllValues = COALESCE(@AllValues + ';', '') + email 
									from cc_email inner join cc_resp on cc_resp.resp = cc_email.id where cc_resp.cc=?
									select @AllValues as email_sent""", owner.Ccusto).fetchval()
			body2 = """<html><head> <font face="Calibri" size="3" color = "black"></head> <body> Alertamos que os gastos do Centro de Custo <b>""" + owner.Ccusto[-3:] + u""" </b> sob sua responsabilidade, excedeu 80% do previsto no Forecast para o ano/mês <b>""" + cdate + """</b> conforme detalhado abaixo: <br><br></font></body></html>"""	
			if sent == None:
				print u"Nao ha RESPONSAVEL PARA O: CC: " + owner.Ccusto 
			else:
				ins_cc = "CC: " + owner.Ccusto
				emailto = str(sent)
				msgfinal = ''.join(body)
				msglala = body2 + msgfinal
				msg = MIMEText(msglala, 'html', 'utf-8')
				msg["From"] = emailfrom
				msg["Subject"] = sub
				msg["To"] = emailto
				server = smtplib.SMTP('mail.maxionwheels.com',25)
				server.starttls()
				try:
					server.sendmail(emailfrom, emailto.split(';'), msg.as_string())
				except SocketError:
					print u"Nao foi possivel conectar ao servidor"
					cursor.execute("INSERT INTO logs VALUES (?, ?, ?, 'Email enviado com sucesso');", ins_date, ins_cc, sent)
					cursor.commit()
					msglala = ""
				else:
					print u"Email enviado com sucesso " + sent
					cursor.execute("INSERT INTO logs VALUES (?, ?, ?, 'Email enviado com sucesso');", ins_date, ins_cc, sent)
					cursor.commit()
					msglala = ""
		else:
			print u"O real nao foi excedido para nenhum grupo dentro do CC: " + str(owner.Ccusto)
			
def conn(ambiente):
	if ambiente == "aco":
		database = parser.get('amb_orcam','database')
		print database
		cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+
					';PWD='+password)
		cnn.setencoding(str, encoding='utf-8')
		cnn.setencoding(unicode, encoding='utf-8', ctype=pyodbc.SQL_CHAR)
		cursor = cnn.cursor()
		app(cnn, cursor)
	if ambiente == "alum":
		database = parser.get('amb_orcam','database')
		cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+
					';PWD='+password)
		cnn.setencoding(str, encoding='utf-8')
		cnn.setencoding(unicode, encoding='utf-8', ctype=pyodbc.SQL_CHAR)
		cursor = cnn.cursor()
		app(cnn, cursor)	
		
def logs(server, username, password, driver, port, ins_date):
	database = parser.get('amb_orcam','database')
	cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+
				';PWD='+password)
	cnn.setencoding(str, encoding='utf-8')
	cnn.setencoding(unicode, encoding='utf-8', ctype=pyodbc.SQL_CHAR)
	cursor = cnn.cursor()
	body = []
	for ins_rel in cursor.execute("select * from logs where data=?",ins_date):
		body.append(u"""<html><head><font face="Calibri" size="3" color = "black">  </head><body> <br>Data:</b> """ + str(ins_rel[0]) + u""" - """ + str(ins_rel[1]) + u""" - <b>Resp:</b> """ + str(ins_rel[2]) + u""" - <b>Mensagem:</b> """ + str(ins_rel[3]) + u"""<br></font></body></html>""")
	msg_rel = ''.join(body)
	msg = MIMEText(msg_rel, 'html', 'utf-8')
	emailto = "rafael.oliveira@maxionwheels.com"
	msg["From"] = emailfrom
	msg["Subject"] = "LOGS ORCAM: CONTA CONTABIL"
	msg["To"] = emailto
	server = smtplib.SMTP('mail.maxionwheels.com',25)
	server.starttls()
	try:
		server.sendmail(emailfrom, emailto, msg.as_string())
	except Exception:
		print u"Não foi possível se conectar ao servidor"
	else:
		print u"LOGS ENVIADO COM SUCESSO"

ambiente = "aco"
conn(ambiente)
logs(server, username, password, driver, port, ins_date)

#ambiente = "alum"
#conn(ambiente)


