import sys
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtSql import *
import pyodbc


server = 'LIM-SQL12QA1'
database = 'TEST_ORCAM_LMS'
username = 'dsserver'
password = 'MaxionHOM123@'
driver = '{SQL Server}' # Driver necessário para conecta ao banco de dados do MSSQL
port = '1433'

# DB type, host, user, password...
db = QSqlDatabase.addDatabase("QODBC");
db.setHostName("LIM-SQL12QA1")
db.setDatabaseName("TEST_ORCAM_LMS")
db.setUserName("dsserver")
db.setPassword("MaxionHOM123@")
ok = db.open()

cnn = pyodbc.connect('DRIVER='+driver+';PORT=port;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+password)
#cnn.setencoding(str, encoding='utf-8')
#cnn.setencoding(unicode, encoding='utf-8', ctype=pyodbc.SQL_CHAR)
cursor = cnn.cursor()

# True if connected


	
row = cursor.execute("select * from ccusto where ccusto >= '0100'")
#for row in cursor.fetchall():


 
class MyTable(QTableWidget):
	def __init__(self, *args):
		QTableWidget.__init__(self, *args)
		self.setmydata()
		self.resizeColumnsToContents()
		self.resizeRowsToContents()
 
	def setmydata(self):
		self.setRowCount(cursor.rowcount)
		for row, form in enumerate(cursor):
			for column, item in enumerate(form):
				self.setItem(row, column, self.QtGui.QTableWidgetItem(str(item)))  
 
def main(args):
    app = QApplication(args)
    table = MyTable(500, 300)
    table.show()
    sys.exit(app.exec_())
 
if __name__=="__main__":
    main(sys.argv)