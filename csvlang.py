#!/usr/bin/python3

import pyexcel as p
import string
import os.path
import re

if not os.path.isfile("config.csv"):
	print("You must have a config.csv file set up.")
	exit()

def err(*messages, fatal=False):
	err_label = "ERROR:" if not fatal else "FATAL:"
	print(err_label, " ".join(messages))
	if fatal:
		print("Exiting due to the last fatal error...")
		exit()

# https://stackoverflow.com/a/12640614
def col2num(col):
	num = 0
	for c in col:
		if c in string.ascii_letters:
			num = num * 26 + (ord(c.upper()) - ord('A')) + 1
	return num-1

# Stands for "String Clean"
# Removes loading and trailing whitespace
def sc(str):
	return str.lstrip().rstrip()

working_input_file = ""
working_output_file = ""
working_sheet = None
config = p.get_sheet(file_name="config.csv")
indexes_to_delete = []

config_sheet = config.get_array()

for row_index,row in enumerate(config_sheet):
	value = sc(row[0].upper())
	if value == "FILE":
		if (len(row) >= 2):
			working_input_file = sc(row[1])
			working_output_file = sc(row[2])
			working_sheet = p.get_sheet(file_name=working_input_file)
			print("FILE:", working_input_file)
		else:
			err("FILE calls must include an input and output file.", fatal=True)
	elif value == "ROW":
		err("Row processing is not implemented yet.")
	elif value == "COL":
		if (working_sheet):
			print("hi")
		else:
			err("You are trying to do a", value, "operation when a sheet is not loaded.", fatal=True)


# for sheet_index,row in enumerate(sheet.row):
# 	for i in range(1, len(config_sheet)):
# 		search_col = col2num(config_sheet[i][0])
# 		search_operation = config_sheet[i][1]
# 		search_str = config_sheet[i][2]

# 		cell_value = row[search_col]

# 		def DeleteIfNeeded(op_condition, should_keep):
# 			if search_operation == op_condition and not should_keep:
# 				print("Deleting row", sheet_index+1)
# 				indexes_to_delete.append(sheet_index)

# 		DeleteIfNeeded("contains" ,   search_str in cell_value)
# 		DeleteIfNeeded("!contains",   search_str not in cell_value)
# 		DeleteIfNeeded("="        ,   search_str == cell_value)
# 		DeleteIfNeeded("!="       ,   search_str != cell_value)
# 		DeleteIfNeeded("regcontains", re.search(search_str, cell_value))
# 		DeleteIfNeeded("!regcontains", not re.search(search_str, cell_value))

# for index in sorted(indexes_to_delete, reverse=True):
# 	del sheet.row[index]

# save_dest = config.row[0][1]
# if (save_dest == "~print~"):
# 	print(sheet)
# else:
#   sheet.save_as(save_dest)
