import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtSql import *
import dbconn
import time


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
		self.progress = QtGui.QProgressBar(self)
		self.progress.setGeometry(200, 80, 250, 20)
		
		#Botão de Progresso
		self.btn = QtGui.QPushButton('Download', self)
		self.btn.move(200, 120)
		self.btn.clicked.connect(self.download)
		
		
		
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
	def download(self):
		dbconn.cursor.execute("delete from orcam")
		dbconn.cursor.commit()
		self.completed = 0
		fileName = 'ORCAM.txt'
		with open('ORCAM.txt') as f:
			line_count = 0
			for line in f:
				line_count += 1
		f.close()
		countline = (100 / line_count)
		print(countline)
		while self.completed <= 100:
			with open(fileName) as f:
				groupList = f.readlines()
				for line in groupList:
					self.progress.setValue(self.completed)
					#print(line)
					#print(line[:4])
					#print(line[4:8])
					#print(line[8:14])
					dbconn.cursor.execute("update dbo.orcam set Vr_Budget=?, Vr_real=? where Ccusto=? and Grupo=? and Anomes=?",line[14:29], line[29:44], line[:4], line[4:8], line[8:14])
					#dbconn.cursor.execute("INSERT INTO orcam VALUES (?,?,?,?,?,1,1,?,?,1,?,NULL,NULL,NULL,NULL)", line[:4], line[4:8], line[8:14], line[14:29], line[29:44], line[14:29], line[14:29], line[14:29])
					dbconn.cursor.commit()
					self.completed += countline
					#print(self.completed)

def run():
	app = QtGui.QApplication(sys.argv)
	GUI = Window()
	GUI.show()
	sys.exit(app.exec_())


if __name__ == '__main__':
	run()  