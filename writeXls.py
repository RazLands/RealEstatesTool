# -*- coding: utf-8 -*-

from xlwt import Workbook
from xlutils.copy import copy
import xlrd, xlwt, os, datetime


def saveWorkSpace(fields, file_path):

	rb = xlrd.open_workbook(file_path, formatting_info=True)
	r_sheet = rb.sheet_by_index(0)
	r = r_sheet.nrows
	wb = copy(rb)
	sheet = wb.get_sheet(0)
	style_text = xlwt.easyxf("align: wrap on, vert centre, horiz centre")

	parameters = [
		datetime.datetime.now().strftime("%d/%m/%Y"),
		fields['status'],
		fields['owner'],
		fields['agent'],
		fields['city'],
		fields['neighborhood'],
		fields['address'],
		fields['type'],
		fields['rooms'],
		fields['floor'],
		fields['area'],
		fields['extra'],
		fields['price'],
		fields['phone'],
		fields['email'],
		fields['comments']
	]

	for i, value in enumerate(parameters):
		# print(value)
		sheet.write(r, i, value, style_text)

	wb.save(file_path)
	return True


def update_excel(row, col, value, file_path):
	# Check if file exists and creates a new one if not
	if not os.path.exists(file_path):
		print(f"File {file_path} wasn't found, Creating a new file")
		create_first_file(file_path)
	rb = xlrd.open_workbook(file_path, formatting_info=True)
	r_sheet = rb.sheet_by_index(0)
	wb = copy(rb)
	sheet = wb.get_sheet(0)
	sheet.write(row, col, value)
	wb.save(file_path)


def create_first_file(file_path):
	wb = Workbook()
	ws = wb.add_sheet("Sheet1")
	header = ["תאריך", "סטטוס", "סוכן", "בעל הנכס", "עיר/ישוב", "שכונה", "כתובת", "סוג נכס", "חדרים", "קומה", 'מ"ר',
	          "תוספות", "מחיר", "מס' טלפון", 'דוא"ל', "הערות", ]
	style_text = xlwt.easyxf("align: vert centre, horiz centre")
	for i, value in enumerate(header):
		ws.col(i).width = 256 * 10
		ws.write(0, i, value, style_text)
	ws.cols_right_to_left = 1
	wb.save(file_path)


def search_excel(keyword, file_path):
	rb = xlrd.open_workbook(file_path, formatting_info=True)
	r_sheet = rb.sheet_by_index(0)
	line_num = -1

	for row in r_sheet.get_rows():
		line_num += 1
		n = len(row)

		if keyword in str(row[n - 1]):
			row.append(line_num)
			yield row
			continue

		backup = row[n - 1]
		row[n - 1] = keyword

		i = 0
		found = 0
		while found == 0 and i < n - 1:

			if keyword in str(row[i]):
				row[n - 1] = backup
				row.append(line_num)
				yield row
				found = 1

			i += 1


# yield count
# while i < n:
#
# 	if keyword in str(row[i]):
#
# 		row[n-1] = backup
# 		if i < n-1:
# 			yield row
#
# 	i += 1

"""
# Data to specified cells.
def writedata(rowNumber, columnNumber, file_path, data):
	book = openpyxl.load_workbook(file_path)
	sheet = book['Sheet1']

	sheet.cell(row=rowNumber, column=columnNumber).value = data
	book.save(file_path)
	print('saved')
"""

"""
def xlrd_open_file(file_path):
	rb = xlrd.open_workbook(file_path, formatting_info=True)
	r_sheet = rb.sheet_by_index(0)
	r = r_sheet.nrows
"""
