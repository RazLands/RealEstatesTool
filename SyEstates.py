# -*- coding: utf-8 -*-
# Version 1.0 1/1/2019
import os, sys
# import subprocess
# subprocess.call("ui2py.bat qt_syestates")

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QMenuBar
from PyQt5.QtWidgets import QLabel, QLineEdit, QComboBox, QTableView, QTableWidget
from PyQt5.QtWidgets import QTableWidgetItem, QStatusBar, QMessageBox, QAction, qApp
from qt_syestates import Ui_MainWindow
from writeXls import saveWorkSpace, search_excel, update_excel, create_first_file

# eel.init("web")
# eel.start("main.html")


class MainWindow(QMainWindow, Ui_MainWindow):
	def __init__(self):
		super(MainWindow, self).__init__()
		self.setupUi(self)
		self.retranslateUi(self)
		_translate = QtCore.QCoreApplication.translate
		self.file_path = 'properties.xls'
		self.changed_values = []
		self.lines = []

		self.excel_but.clicked.connect(self.write_to_excel)
		self.search_but.clicked.connect(self.search)
		self.clear_res_but.clicked.connect(self.clear_table)
		self.update_but.clicked.connect(self.update_excel_clicked)

		self.menuBar_actions()

	def menuBar_actions(self):
		#Open excel
		self.openXL.setStatusTip("Open {}".format(self.file_path))
		# self.openXL.triggered.connect(os.system('start "%s\\properties.xls"' % (sys.path[0])))

		#Exit app
		self.exit.setStatusTip('Exit application')
		self.exit.triggered.connect(qApp.quit)

	def open_XL(self):
		os.chdir(sys.path[0])
		os.system('start properties.xls')
		# os.system('start "%s\\properties.xls"' % path)

	def cell_clicked(self, item):
		print("Changed values1: ", self.changed_values)
		row = item.row()
		col = item.column()
		data = self.result_table.item(row, col).text()
		self.changed_values.append((row, col, data))
		# print("Value changed")
		print("Changed values2: ", self.changed_values)
		print("Changed data: ", self.changed_values[0][2])
		print(data)

	def update_excel_clicked(self):
		# self.statusbar.showMessage("Checking this works")
		print("Updating new values: ", self.changed_values)
		for row, col, value in self.changed_values:
			update_excel(self.lines[row], col, value, self.file_path)

		self.changed_values = []

	def write_to_excel(self):
		# Check if file exists and creates a new one if not
		if not os.path.exists(self.file_path):
			QMessageBox.about(self, "Create New File", "קובץ לא קיים.\nלחץ אישור ליצירת קובץ חדש")
			create_first_file(self.file_path)

		agent = self.agent_combox.currentText()
		owner = self.owner_edit.text()
		phone = self.phone_edit.text()
		email = self.email_edit.text()
		address = self.strtNum_edit.text()
		neighborhood = self.neighborhood_edit.text()
		city = self.city_edit.text()
		proptype = self.type_combox.currentText()
		area = self.area_edit.text()
		rooms = self.rooms_edit.text()
		floor = self.floorNum_edit.text()
		price = self.price_edit.text()
		extra = {
			"lift": "מעלית\n" if self.lift_box.isChecked() else "",
			"park": "חניה\n" if self.park_box.isChecked() else "",
			"shelter": "ממד\n" if self.shelter_box.isChecked() else "",
			"renew": "משופצת\n" if self.renew_box.isChecked() else "",
			"equipped": "מרוהטת\n" if self.equipped_box.isChecked() else "",
			"balcony": "מרפסת\n" if self.balcony_box.isChecked() else ""
		}
		comments = self.comments_edit.text()
		fields = {
			"agent": agent,
			"owner": owner,
			"phone": phone,
			"email": email,
			"address": address,
			"neighborhood": neighborhood,
			"city": city,
			"type": proptype,
			"area": area,
			"rooms": rooms,
			"floor": floor,
			"price": price,
			"extra": extra["lift"] + extra["park"] + extra["shelter"] + extra["renew"] + extra["equipped"] + extra[
				"balcony"],
			"comments": comments,
			"status": "מודעה פעילה"
		}
		# print(fields)
		try:
			if saveWorkSpace(fields, self.file_path):
				self.statusbar.showMessage(f"Property owned by {fields['owner']} have been added to the file {self.file_path}")
				QMessageBox.about(self, "Done!", f"הנכס בבעלות {fields['owner']} הוכנס לקובץ בהצלחה!")
		except IOError:
			self.statusbar.showMessage(f"The file {self.file_path} need to be closed before continue!")
			QMessageBox.about(self, "שגיאה!", f"שגיאה!\nיש לסגור את הקובץ {self.file_path} לפני ביצוע הפעולה!")

	def search(self):
		self.lines = []
		self.result_table.setRowCount(0)
		keyword = str(self.search_edit.text()).strip(",")
		result = search_excel(keyword, self.file_path)

		for row in result:
			n = len(row)
			self.lines.append(row.pop(n - 1))
			rowPos = self.result_table.rowCount()
			self.result_table.insertRow(rowPos)
			# print("rowPos: ", rowPos)

			for i, cell in enumerate(row):
				# print("index: ", i, "cell.value: ", cell.value)
				self.result_table.setItem(rowPos, i, QtWidgets.QTableWidgetItem(str(cell.value)))
		self.result_table.itemChanged.connect(self.cell_clicked)
		self.changed_values = []
		self.statusbar.showMessage(f"Showing search results for: {keyword}")

	def clear_table(self):
		self.result_table.clearContents()
		self.result_table.setRowCount(0)
		self.lines = []
		self.changed_values = []


app = QApplication(sys.argv)
ui = MainWindow()
ui.show()
sys.exit(app.exec_())
