# encoding: utf-8
# encoding: iso-8859-1
# encoding: win-1252
import pyodbc
import time

cdate = time.strftime("%Y%m")
mes = time.strftime("%m")
cmonth = int(mes) + 1

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
#print "Escolha uma das opções abaixo:"
#ambiente = raw_input("- [1] - Aço\n- [2] - Aluminio\n")
#os.system('cls')
#file = raw_input("")


with open("CargaFCSTACO103.csv", "r") as filestream:
	for line in filestream:
		currentline = line.split(",")
		cc = '0'+str(currentline[0])
		grupo = '0'+str(currentline[1][-3:])
		#grupo = str(currentline[1])
		forecast = float(currentline[cmonth])
		try:
			#print currentline
			print "Paramentro: 0 CC: " + cc + " Grupo: " + grupo + " Ano/Mes: " + cdate + " Forecast: " + str(forecast)
			print "\n"
		except Exception:
			print "Checar caracter"
			#for lala in cursor.execute("select * from orcam where Ccusto=? and grupo=? and anomes=?", cc, grupo, cdate):
			#	query = "update dbo.orcam set Tx_Forecast=0, Vr_Forecast=" + str(forecast) + " where Ccusto=" + cc + " and Grupo=" + grupo + " and Anomes=" + cdate +""
			#	print query
			#	cursor.execute(query)
		cursor.execute("update dbo.orcam set Tx_Forecast=0, Vr_Forecast=? where Ccusto=? and Grupo=? and Anomes=?", forecast, cc, grupo, cdate)
		cursor.commit()
			