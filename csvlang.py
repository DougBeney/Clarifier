import pyexcel as p
import string
import os.path


if not os.path.isfile("config.csv"):
	print("You must have a config.csv file set up.")
	exit()

# https://stackoverflow.com/questions/7261936/convert-an-excel-or-spreadsheet-column-letter-to-its-number-in-pythonic-fashion/12640614
def col2num(col):
	num = 0
	for c in col:
		if c in string.ascii_letters:
			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
	return num-1

config = p.get_sheet(file_name="config.csv")
sheet = p.get_sheet(file_name=config.row[0][0])

for sheet_index,row in enumerate(sheet.row):
	config_sheet = config.get_array()
	for i in range(1, len(config_sheet)):
		search_col = col2num(config_sheet[i][0])
		search_operation = config_sheet[i][1]
		search_str = config_sheet[i][2]
		if search_operation == "contains":
			if search_str not in row[search_col] :
				print("Deleting row", sheet_index+1)
				del sheet.row[sheet_index]
		elif search_operation == "!contains":
			if search_str in row[search_col]:
				print("Deleting row", sheet_index+1)
				del sheet.row[sheet_index]
		elif search_operation == "=":
			if search_str == row[search_col]:
				print("Deleting row", sheet_index+1)
				del sheet.row[sheet_index]
		elif search_operation == "!=":
			if search_str != row[search_col]:
				print("Deleting row", sheet_index+1)
				del sheet.row[sheet_index]

sheet.save_as(config.row[0][1])
