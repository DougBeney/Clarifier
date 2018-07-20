#!/usr/bin/python3

import pyexcel as p
import string
import os.path
import re

working_input_file = ""
working_output_file = ""
working_sheet = None
working_line_number = 0
working_row_indexes_to_delete = []
working_col_indexes_to_delete = []

if not os.path.isfile("config.csv"):
	print("You must have a config.csv file set up.")
	exit()

config = p.get_sheet(file_name="config.csv")
config_sheet = config.get_array()

def err(*messages, fatal=False):
	err_label = "ERROR:" if not fatal else "FATAL:"
	print(err_label, " ".join(messages), "[Line #"+str(working_line_number)+"]")
	if fatal:
		print("Exiting due to the last fatal error...") exit()
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

def execute_and_save():
	for index in sorted(working_col_indexes_to_delete, reverse=True):
		del working_sheet.row[index]
	for index in sorted(working_row_indexes_to_delete, reverse=True):
		del working_sheet.col[index]
	if (working_output_file == "~print~"):
		print(working_sheet)
	else:
		working_sheet.save_as(working_output_file)
	working_row_indexes_to_delete = []
	working_col_indexes_to_delete = []

for row_index,row in enumerate(config_sheet):
	OPERATION = sc(row[0].upper())
	working_line_number = row_index + 1
	if OPERATION == "FILE":
		if (len(row) >= 2):
			if working_sheet is not None:
				execute_and_save()
			working_input_file = sc(row[1])
			working_output_file = sc(row[2])
			working_sheet = p.get_sheet(file_name=working_input_file)
			print("FILE:", working_input_file)
		else:
			err("FILE calls must include an input and output file.", fatal=True)
	elif working_sheet:
		elif OPERATION == "ROW" or OPERATION == "COL":
			# 0        1    2         3
			# ROW/COL, 1/a, operator, search
			if (len(row) < 4):
				err(OPERATION, "operation requires 3 additional arguments", fatal=True)
			sheet_array         = None
			location            = row[1]
			search_operator     = row[2]
			search_str          = row[3]
			working_sheet_index = 0

			if   OPERATION == "ROW":
				# get col array so we can easily check a row
				sheet_array = working_sheet.col
			elif OPERATION == "COL":
				# get row array so we can easily check a col
				sheet_array = working_sheet.row
				# must convert col letter to number
				location = col2num(location)

			def DeleteIfNeeded(op_condition, should_keep):
				global OPERATION
				if search_operator == op_condition and not should_keep:
					print("Deleting", OPERATION, working_sheet_index+1)
					if OPERATION == "ROW":
						working_row_indexes_to_delete.append([OPERATION, working_sheet_index])
					elif OPERATION == "COL":
						working_col_indexes_to_delete.append([OPERATION, working_sheet_index])

			for sheet_index,row_or_col in enumerate(working_sheet.row):
				cell_value = row_or_col[location]
				working_sheet_index = sheet_index
				DeleteIfNeeded("contains"    , search_str in cell_value)
				DeleteIfNeeded("!contains"   , search_str not in cell_value)
				DeleteIfNeeded("="           , search_str == cell_value)
				DeleteIfNeeded("!="          , search_str != cell_value)
				DeleteIfNeeded("regcontains" , re.search(search_str, cell_value))
				DeleteIfNeeded("!regcontains", not re.search(search_str, cell_value))
	else
		err("You are trying to do a", value, "operation when a sheet is not loaded.", fatal=True)

execute_and_save()

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
