# encoding: utf-8
# encoding: iso-8859-1
# encoding: win-1252
import pyodbc
import sys
import os
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import smtplib
import mimetypes
import shutil
import subprocess
import time
from distutils.core import setup

server = 'LIM-SQL12P1'
username = 'dsserver'
password = 'Maxion123@'
driver = '{SQL Server}' # Driver necessário para conecta ao banco de dados do MSSQL
port = '1433'

#Função para definir o ambiente: aço ou aluminio
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
	quest(cursor)
		
#Principal função do sistema onde ocorre toda interação
def quest(cursor):
	os.system('cls') 
	print "*******************************************************************"
	print u"Olá, seja bem-vindo."
	print u"\nDesenvolvido por:"
	print u" - Rafael Oliveira"
	print u" - Contato: rafael.oliveira@maxionwheels.com - (19) 3404-2360"
	print "*******************************************************************\n"
	quest01 = raw_input(u"Por favor, escolha uma opcao:\n\nCADASTRAR...\n- [1] - cadastrar um novo responsavel para um Centro de Custo;\n- [2] - cadastrar um novo responsavel para uma Conta Conatbil;\n\nREMOVER...\n- [3] - remover um responsavel por um CC especifico;\n- [4] - remover um responsavel por uma Conta Contabil especifica.\n\nLISTAR... \n- [5] - Listar responsavel por um CC especifico.\n- [6] - Listar responsavel por uma Conta especifica\n")
	if quest01 == '1':
		os.system('cls') 
		print u"Voce escolheu a opcao 1, portanto, vamos cadastrar um novo responsavel para Centro de Custo\n"
		quest02 = raw_input(u"Digite para qual Centro de Custo gostaria de cadastrar um responsavel. Colocar '0' antes do centro de custo. \n - Exemplo: 0151\n")
		cursor.execute("select * from Ccusto where Ccusto=?", quest02)
		ver_cc = cursor.fetchone()
		if ver_cc == None:
			print u"Voce digitou um CC de custo errado. Por favor, recomece!"
			print "Finalizando.."
			time.sleep(3)
			sys.exit()
		else:
			quest03 = raw_input(u"Agora digite o e-mail do responsavel: \n")
			cursor.execute("select id, email from cc_email where email=?", quest03)
			ver_email = cursor.fetchone()
			if ver_email == None:
				cursor.execute("INSERT INTO cc_email VALUES (?);", quest03)
				cursor.commit()
				cursor.execute("select * from cc_email where email=?", quest03)
				get_id_email = cursor.fetchone()
				cursor.execute("INSERT INTO cc_resp VALUES (?, ?);", quest02, get_id_email.id)
				cursor.commit()
				print u"\nO e-mail: " + quest03 + " foi adicionado e associado ao CC: " + quest02 + "."
				result_check = "ok"
			else:
				cursor.execute("select * from cc_email where email = ?", quest03)
				get_id_email = cursor.fetchone()
				get_id = get_id_email.id
				result_check = check(cursor, quest02, get_id, quest01)
			if result_check == None:
				cursor.execute("INSERT INTO cc_resp VALUES (?, ?);", quest02, get_id_email.id)
				cursor.commit()
				print "\n*******************************************************************"
				print u"O email foi adicionado com sucesso ao CC: " + quest02
				print "*******************************************************************\n"
				time.sleep(5)
	elif quest01 == '2':
		os.system('cls') 
		print u"Voce escolheu a opcao 2, portanto, vamos cadastrar um novo responsavel para Conta Contabil\n"
		quest02 = raw_input(u"Digite para qual Conta Contabil gostaria de cadastrar um responsavel. Colocar '0' antes da Conta. \nExemplo: 0151\n")
		cursor.execute("select * from Grupos where Grupo=?", quest02)
		ver_cc = cursor.fetchone()
		if ver_cc == None:
			print u"Voce digitou uma Conta Contábil errado. Por favor, recomece!"
			print "Finalizando.."
			time.sleep(3)
			sys.exit()
		else:
			quest03 = raw_input(u"Agora digite o e-mail do responsavel: \n")
			cursor.execute("select id, email from grupo_email where email=?", quest03)
			ver_email = cursor.fetchone()
			if ver_email == None:
				cursor.execute("INSERT INTO grupo_email VALUES (?);", quest03)
				cursor.commit()
				cursor.execute("select * from grupo_email where email=?", quest03)
				get_id_email = cursor.fetchone()
				cursor.execute("INSERT INTO grupo_resp VALUES (?, ?);", quest02, get_id_email.id)
				cursor.commit()
				print "\n*******************************************************************\n"		
				print u"O e-mail: " + quest03 + " foi adicionado e associado a Conta Conatabil: " + quest02 + "."
				print "\n*******************************************************************\n"		
				result_check = "ok"
			else:
				cursor.execute("select * from grupo_email where email = ?", quest03)
				get_id_email = cursor.fetchone()
				get_id = get_id_email.id
				result_check = check(cursor, quest02, get_id, quest01)
			if result_check == None:
				cursor.execute("INSERT INTO grupo_resp VALUES (?, ?);", quest02, get_id_email.id)
				cursor.commit()
				print "\n*******************************************************************\n"
				print u"O email foi adicionado com sucesso a Conta Contabil: " + quest02
				print "\n*******************************************************************\n"			
	elif quest01 == '3':
		os.system('cls')
		print u"Voce escolheu a opcao 3, portanto, vamos remover um responsável associado a um CC\n"
		quest02 = raw_input(u"Digite para qual CC gostaria de remover um responsavel. Colocar '0' antes do CC. \nExemplo: 0151\n")
		for listall in cursor.execute("select cc_resp.cc, cc_email.id, cc_email.email from cc_resp inner join cc_email on cc_resp.resp = cc_email.id where cc_resp.cc =? order by cc_email.id", quest02):
			print u" - CC: " + str(listall[0]) + " ID: " + str(listall[1]) + u" E-mail: " + str(listall[2])
		cursor.execute("select * from Ccusto where Ccusto=?", quest02)
		ver_cc = cursor.fetchone()
		if ver_cc == None:
			print u"Voce digitou um CC errado. Por favor, recomece!"
			print "Finalizando.."
			time.sleep(3)
			sys.exit()
		else:
			quest03 = raw_input(u"\nAgora, digite o numero do ID que deseja remover relacionado ao responsavel: \n")
			cursor.execute("delete from cc_resp where cc=? and resp=?", quest02, quest03)
			cursor.commit()
			os.system('cls')
			print "*******************************************************************"
			print u"OK - responsavel removido com sucesso. Para ter certeza, vou gerar novamente a listagem do responsável pelo CC: " + quest02
			print "*******************************************************************"
			time.sleep(3)
			os.system('cls')
			for listall in cursor.execute("select cc_resp.cc, cc_email.id, cc_email.email from cc_resp inner join cc_email on cc_resp.resp = cc_email.id where cc_resp.cc =? order by cc_email.id", quest02):
				print u" - CC: " + str(listall[0]) + " ID: " + str(listall[1]) + u" E-mail: " + str(listall[2])
	elif quest01 == '4':
		os.system('cls')
		print u"Voce escolheu a opcao 4, portanto, vamos remover um responsável associado a uma Conta Contabil\n"
		quest02 = raw_input(u"Digite para qual Conta gostaria de remover um responsavel. Colocar '0' antes da Conta. \nExemplo: 0411\n")
		for listall in cursor.execute("select grupo_resp.grupo, grupo_email.id, grupo_email.email from grupo_resp inner join grupo_email on grupo_resp.resp = grupo_email.id where grupo_resp.grupo =? order by grupo_email.id", quest02):
			print u" - Conta: " + str(listall[0]) + " ID: " + str(listall[1]) + u" E-mail: " + str(listall[2])
		cursor.execute("select * from Grupos where Grupo=?", quest02)
		ver_cc = cursor.fetchone()
		if ver_cc == None:
			print u"Voce digitou uma Conta errado. Por favor, recomece!"
			print "Finalizando.."
			time.sleep(3)
			sys.exit()
		else:
			quest03 = raw_input(u"\nAgora, digite o numero do ID que deseja remover relacionado ao responsavel: \n")
			cursor.execute("delete from grupo_resp where grupo=? and resp=?", quest02, quest03)
			cursor.commit()
			os.system('cls')
			print "*******************************************************************"
			print u"OK - responsavel removido com sucesso. Para ter certeza, vou gerar novamente a listagem do responsável pelo Conta: " + quest02
			print "*******************************************************************"
			time.sleep(3)
			os.system('cls')
			for listall in cursor.execute("select grupo_resp.grupo, grupo_email.id, grupo_email.email from grupo_resp inner join grupo_email on grupo_resp.resp = grupo_email.id where grupo_resp.grupo =? order by grupo_email.id", quest02):
				print u" - CC: " + str(listall[0]) + " ID: " + str(listall[1]) + u" E-mail: " + str(listall[2])
	elif quest01 == '5':
		os.system('cls')
		print u"Voce escolheu a opcao 5, portanto, vamos listar os responsaveis para o CC\n"
		quest02 = raw_input(u"Digite para qual CC gostaria de listar os responsaveis cadastrado. Colocar '0' antes do CC. \nExemplo: 0151\n")
		os.system('cls')
		print "\n\n*******************************************************************"
		for listall in cursor.execute("select cc_resp.cc, cc_email.id, cc_email.email from cc_resp inner join cc_email on cc_resp.resp = cc_email.id where cc_resp.cc =? order by cc_email.id", quest02):
			print u" - CC: " + str(listall[0]) + " ID: " + str(listall[1]) + u" E-mail: " + str(listall[2])
		print "*******************************************************************"
		time.sleep(5)	
	elif quest01 == '6':
		os.system('cls')
		print u"Voce escolheu a opcao 6, portanto, vamos listar os responsaveis para uma Conta Contabil\n"
		quest02 = raw_input(u"Digite para qual Conta gostaria de listar os responsaveis cadastrado. Colocar '0' antes do Grupo. \nExemplo: 0411\n")
		os.system('cls')
		print "\n\n*******************************************************************"
		for listall in cursor.execute("select grupo_resp.grupo, grupo_email.id, grupo_email.email from grupo_resp inner join grupo_email on grupo_resp.resp = grupo_email.id where grupo_resp.grupo =? order by grupo_email.id", quest02):
			print u" - Conta: " + str(listall[0]) + " ID: " + str(listall[1]) + u" E-mail: " + str(listall[2])
		print "*******************************************************************"
		time.sleep(5)	
	else:
		print u"Você digitou uma opcao invalida. Por favor, recomece!"
		time.sleep(3)
		sys.exit()
		
#Função chamada para continuar no sistema sem a necessidade de fechar			
def ask():
	newline = raw_input("\nGostaria de utilizar o sistema novamente?\n - [s] - para sim;\n - [n] - para nao. \n")
	return newline

#Função para checar se uma pessoa já está cadastrado num CC ou Conta especificada na entrada de dados	
def check(cursor, quest02, get_id, quest01):
	if quest01 == '1':
		cursor.execute("select * from cc_resp where cc=? and resp=?", quest02, get_id)
		chk_cc_gr = cursor.fetchone()
		if chk_cc_gr == None:
			return chk_cc_gr
		else:
			print "\n*******************************************************************\n"
			print u"\nO responsavel já está cadastrado no CC mencionado.\n"
			print "\n*******************************************************************\n"
			return chk_cc_gr
			time.sleep(3)
	elif quest01 == '2':
		cursor.execute("select * from grupo_resp where grupo=? and resp=?", quest02, get_id)
		chk_cc_gr = cursor.fetchone()
		if chk_cc_gr == None:
			return chk_cc_gr
		else:
			print "\n*******************************************************************\n"
			print u"\nO responsavel já está cadastrado na Conta Contabil mencionado.\n"
			print "\n*******************************************************************\n"
			return chk_cc_gr
			time.sleep(3)

#Chama a função ambiente. Começo do programa...
ambiente()
#Chama a função ask com o objetivo de perguntar ao usuário se ele quer continuar ou não com o programa
yn = ask()
while (yn == 's'):
	ambiente()
	yn = ask()
	if (yn == 'n'):
		print u"Obrigado por utilizar... fechando...!"
		time.sleep(3)
		break
	else:
		True