#!/usr/bin/python3

import pyexcel as p
import string
import os.path
import re

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
indexes_to_delete = []

for sheet_index,row in enumerate(sheet.row):
	config_sheet = config.get_array()
	for i in range(1, len(config_sheet)):
		search_col = col2num(config_sheet[i][0])
		search_operation = config_sheet[i][1]
		search_str = config_sheet[i][2]

		cell_value = row[search_col]

		def DeleteIfNeeded(op_condition, should_keep):
			if search_operation == op_condition and not should_keep:
				print("Deleting row", sheet_index+1)
				indexes_to_delete.append(sheet_index)

		DeleteIfNeeded("contains" ,   search_str in cell_value)
		DeleteIfNeeded("!contains",   search_str not in cell_value)
		DeleteIfNeeded("="        ,   search_str == cell_value)
		DeleteIfNeeded("!="       ,   search_str != cell_value)
		DeleteIfNeeded("regcontains", re.search(search_str, cell_value))
		DeleteIfNeeded("!regcontains", not re.search(search_str, cell_value))

for index in sorted(indexes_to_delete, reverse=True):
	del sheet.row[index]

save_dest = config.row[0][1]
if (save_dest == "~print~"):
	print(sheet)
else:
  sheet.save_as(save_dest)
