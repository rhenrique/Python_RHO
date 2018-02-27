import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtSql import *
import time
import dbconn
import datetime
from PyQt4.QtCore import QThread, pyqtSignal


TIME_LIMIT = 100

class External(QThread):
	countChanged = pyqtSignal(int)
	
	def run(self):
# Importação do MATERS.txt
		self.completed = 0
		fileName = 'MATERS.txt'
		with open(fileName) as f:
			line_count = 0
			line = f.readline()
			for line in f:
				line_count += 1
		f.close()
		countline = (100 / line_count)
		dbconn.cursor.execute("delete from marteriais where anomes=?",line[:6])
		dbconn.cursor.commit()
		NULL = None
		while self.completed <= TIME_LIMIT:
			with open(fileName) as f:
				groupList = f.readlines()
				for line in groupList:
					# # print(line[:6]) #Anomes
					# # print(line[6:21]) #Conta
					# # print(line[21:34]) #Processo
					# # print(line[34:38]) #Ccusto
					dt = datetime.datetime.strptime(line[51:59], '%d%m%Y').date()
					# # print(dt) #Dt_mov
					# # print(line[59:62]) #Codmov
					# # print(line[38:51]) #Codmat
					# # print(line[62:66]) #Grupo
					# # print(line[66:72]) #Codfor
					# # print(line[99:100]) #D/C
					if line[99:100] == 'C':
						dc = line[99:100]
						qt_mov = (float(line[72:84]) / 1000) * (-1)
						vr_mov = (float(line[84:99]) / 100) * (-1)
						# # print(float(line[72:84]))
						# # print(qt_mov)
						# # print(vr_mov)
					else:
						dc = line[99:100]
						qt_mov = float(line[72:84]) / 1000
						vr_mov = float(line[84:99]) / 100
						#print(float(line[72:84]))
						# # print(qt_mov)
						# # print(vr_mov)
					# # print(line[114:115]) #Tipo de Suprimento
					# # print(line[108:114]) #Chapa
					# # print(line[115:121]) #Chapa_ret
					# # print(line[121:122]) #CodApl
					# # print(line[122:135]) #aplica
					# # print(line[100:104]) #setor_rq
					# # print(line[104:108]) #setor_be
					# # print(line[135:148]) #nm_fantasia
					# # print(line[148:198]) #descrição
					# # print(line[198:208]) #eam
					# # print("\n")
					dbconn.cursor.execute("INSERT INTO marteriais VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", 
					line[:6], line[6:21], line[21:34], line[34:38], line[38:51],dt,line[59:62],line[62:66],
					line[66:72],qt_mov,vr_mov,dc, line[108:114],line[114:115],NULL,NULL,NULL,
					NULL,NULL,line[104:108],NULL,NULL,line[115:121],line[121:122],
					line[122:135],line[100:104],line[135:148],line[148:198],line[198:208])
					dbconn.cursor.commit()
					self.completed += countline
					self.countChanged.emit(self.completed)
					
# Importação do LANCAM.txt
					
		self.completed = 0
		fileName = 'LANCAM.txt'
		with open(fileName) as f:
			line_count = 0
			line = f.readline()
			for line in f:
				line_count += 1
		f.close()
		countline = (100 / line_count)
		dbconn.cursor.execute("delete from lancam where anomes=?",line[:6])
		dbconn.cursor.commit()
		NULL = None		
		while self.completed <= TIME_LIMIT:
			with open(fileName) as f:
				groupList = f.readlines()
				for line in groupList:
					print(line[:6]) #anomes
					print(line[6:21]) #conta
					print(line[21:26]) #lancam
					print(line[26:27]) #dc
					print(line[27:31]) #ccusto
					print(line[31:35]) #grupo
					dt = datetime.datetime.strptime(line[35:43], '%d%m%Y').date()
					print(dt) #Dt_lancam
					if line[21:26] == 'C':
						vr_lancam = (float(line[43:58]) / 100) * (-1)
					else:
						vr_lancam = float(line[43:58]) / 100
					print(line[58:93]) #historico
					print(line[101:109]) #usuario
					dbconn.cursor.execute("INSERT INTO lancam VALUES (?,?,?,?,?,?,?,?,?,?)",
					line[:6],line[6:21],line[21:26],line[26:27],line[27:31],line[31:35],dt,
					vr_lancam,line[58:93],line[101:109])
					dbconn.cursor.commit()
					self.completed += countline
					self.countChanged.emit(self.completed)

# # Importação do CCUSTO.txt
					
		self.completed = 0		
		fileName = 'CCUSTO.txt'
		with open(fileName) as f:
			line_count = 0
			for line in f:
				line_count += 1
		f.close()
		countline = (100 / line_count)
		print(countline)
		
		while self.completed <= TIME_LIMIT:
			with open(fileName) as f:
				groupList = f.readlines()
				for line in groupList:
					dbconn.cursor.execute("select * from ccusto where ccusto=?",line[:4])
					chk = dbconn.cursor.fetchone()
					if chk == None:
						print(line[:4])
						dbconn.cursor.execute("INSERT INTO ccusto VALUES (?)", line[:4])
						dbconn.cursor.commit()
					self.completed += countline
					self.progress.setValue(self.completed)
					
# # Importação do ORCAM.txt					
					
		self.completed = 0
		fileName = 'ORCAM.txt'
		with open('ORCAM.txt') as f:
			line_count = 0
			for line in f:
				line_count += 1
		f.close()
		countline = (100 / line_count)
		print(countline)
		while self.completed <= TIME_LIMIT:
			with open(fileName) as f:
				groupList = f.readlines()
				for line in groupList:
					#print(line)
					#print(line[:4])
					#print(line[4:8])
					#print(line[8:14])
					#a = dbconn.cursor.execute("update dbo.orcam set Vr_Budget=?, Vr_real=? where Ccusto=? and Grupo=? and Anomes=?",line[14:29], line[29:44], line[:4], line[4:8], line[8:14])
					dbconn.cursor.execute("select * from orcam where anomes=? and Ccusto=? and Grupo=?",line[8:14], line[:4], line[4:8])
					chk = dbconn.cursor.fetchone()
					#print(str(chk))
					#print(float(line[29:44])) 
					#print(chk[1])
					r = float(line[29:44]) / 100 #Valor Real
					b = float(line[14:29]) / 100 #Valor do AOP
					if chk != None:
						if float(line[29:44]) != float(chk.Vr_real) and chk.Tx_Forecast == 1:
							print(line[4:8])
							print(line[:4])
							print(r)
							print(b)
							#print(chk.Vr_real)
							#print(chk.Ccusto)
							#print(chk.Grupo)
							print('\n')
							dbconn.cursor.execute("update dbo.orcam set Vr_Budget=?, Vr_Forecast=?, Vr_orcado=?, Vr_Forecastu=?, Vr_real=? where Ccusto=? and Grupo=? and Anomes=?",b, b, b, b, r, line[:4], line[4:8], line[8:14])
							dbconn.cursor.commit()
						if float(line[29:44]) != float(chk.Vr_real) and chk.Tx_Forecast == 0:
							dbconn.cursor.execute("update dbo.orcam set Vr_Budget=?, Vr_orcado=?, Vr_Forecastu=?, Vr_real=? where Ccusto=? and Grupo=? and Anomes=?",b, b, b, b, r, line[:4], line[4:8], line[8:14])
							dbconn.cursor.commit()
					else:
						dbconn.cursor.execute("INSERT INTO orcam VALUES (?,?,?,?,?,1,1,?,?,1,?,NULL,NULL,NULL,NULL)", line[:4], line[4:8], line[8:14], b, r, b, b, b)
						dbconn.cursor.commit()
					self.completed += countline
					self.progress.setValue(self.completed)

# Importação do CCUSTOB.txt					
		
		self.completed = 0
		fileName = 'CCUSTOB.txt'
		with open(fileName) as f:
			line_count = 0
			for line in f:
				line_count += 1
		f.close()
		countline = (100 / line_count)
		print(countline)
		
		while self.completed <= TIME_LIMIT:
			with open(fileName) as f:
				groupList = f.readlines()
				for line in groupList:
					dbconn.cursor.execute("select * from ccusto where ccustob=?",line[:4])
					chk = dbconn.cursor.fetchone()
					if chk == None:
						print(line[:4])
						dbconn.cursor.execute("INSERT INTO ccustob VALUES (?,?,'')", line[:4, line[4:44]])
						dbconn.cursor.commit()
					self.completed += countline
					self.progress.setValue(self.completed)	
					

class SecondWindow(QtGui.QMainWindow):
	def __init__(self, parent=None):
		super(SecondWindow, self).__init__(parent)
		self.setWindowTitle("Second Window")
		self.setGeometry(500, 500, 500, 300)
		self.setWindowIcon(QtGui.QIcon('pythonlogo.png'))

class Window(QtGui.QMainWindow):

	def __init__(self, parent=None):
		super(Window, self).__init__(parent)
		self.setWindowTitle("PyQT tuts!")
		self.setGeometry(500, 500, 500, 300)
		self.setWindowIcon(QtGui.QIcon('pythonlogo.png'))

		extractAction = QtGui.QAction("&Forecast", self)
		#extractAction.setShortcut("Ctrl+Q")
		#extractAction.setStatusTip('Leave The App')
		extractAction.triggered.connect(self.close_application)
		
		openAction = QtGui.QAction("&Usuários!!!", self)
		openAction.setShortcut("Ctrl+Q")
		openAction.setStatusTip('Leave The App')
		openAction.triggered.connect(self.close_application)
		
		
		forecast = QtGui.QAction("&Forecast", self)
		budget = QtGui.QAction("&Budget", self)
		
		#tools = fileMenu.addMenu('&Tools')
		#prevMenu = QtGui.QMenu('Preview',self)
		#prevInNuke = QtGui.QAction("Using &Nuke",prevAction)
		#tools.addMenu(prevMenu)
		#prevAction.addAction(prevInNuke)
		
		#mainMenu = self.menuBar()
		#fileMenu = mainMenu.addMenu('&File')
		#fileMenu.addAction(extractAction)
		
		#Progresso
		inMsg = 'Loading..'
		self.progress = QtGui.QProgressBar(self)
		self.progress.setGeometry(200, 80, 250, 20)
		
		#Botão de Progresso
		self.btn = QtGui.QPushButton('Importar ORÇAM', self)
		
		self.btn.move(200, 120)
		#self.btn.clicked.connect(self.materstxt)
		self.btn.clicked.connect(self.onButtonClick)
		
		
		
		#Arquivo
		self.statusBar()
		mainMenu = self.menuBar()
		
		#Menu "Arquivo"
		fileMenu = mainMenu.addMenu('&Arquivo')
		
		#Menu "Usuários"
		users = fileMenu.addMenu('&Novo')
		prevMenu = users.addAction(openAction)
		teste = users.addMenu("Sub menu")
		testeAction = users.addAction(extractAction)
		
		#Menu "Abrir"
		abrir = fileMenu.addMenu('&Abrir')
		prevMenu = abrir.addAction(budget)
		self.home()
		
	def onButtonClick(self):
		self.calc = External()
		self.calc.countChanged.connect(self.onCountChanged)
		self.calc.start()

	def onCountChanged(self, value):
		self.progress.setValue(value)
		
	def on_pushButton_clicked(self):
		self.dialog.show()
		

	def home(self):
		btn2 = QtGui.QPushButton("Second Window", self)
		btn2.resize(btn2.minimumSizeHint())
		btn2.move(0,50)
		#self.show()
		btn2.clicked.connect(self.on_pushButton_clicked)
		self.dialog = SecondWindow(self)
		
		
		btn = QtGui.QPushButton("Quit", self)
		btn.clicked.connect(self.close_application)
		btn.resize(btn.minimumSizeHint())
		btn.move(0,100)
		self.show()

	def close_application(self):
		print("whooaaaa so custom!!!")
		sys.exit()
					
				
def run():
	app = QtGui.QApplication(sys.argv)
	GUI = Window()
	GUI.show()
	sys.exit(app.exec_())


if __name__ == '__main__':
	run()  