import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtSql import *
import dbconn


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


def run():
	app = QtGui.QApplication(sys.argv)
	GUI = Window()
	GUI.show()
	sys.exit(app.exec_())


if __name__ == '__main__':
	run()  