#!/usr/bin/python3
import pyexcel as p
import string
import os.path
import re

working_input_file = ""
working_output_file = ""
working_sheet = None
working_line_number = 0
working_cell_coord = [0,0]
working_row_indexes_to_delete = []
working_col_indexes_to_delete = []

if not os.path.isfile("config.csv"):
	print("You must have a config.csv file set up.")
	exit()


config = p.get_sheet(file_name="config.csv")
config_sheet = config.get_array()


def err(*messages, fatal=False):
	err_label  = "ERROR:" if not fatal else "FATAL:"
	linum_str  = "[Line #"+str(working_line_number)+" in config.csv]"
	cell_linum_str = "[Sheet Coord: "+num2col(working_cell_coord[0]+1)+str(working_cell_coord[1]+1)+"]"
	print(err_label, linum_str, cell_linum_str, " ".join(messages))
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


# https://stackoverflow.com/a/23862195
def num2col(num):
	string = ""
	while num > 0:
		num, remainder = divmod(num - 1, 26)
		string = chr(65 + remainder) + string
	return string


# https://stackoverflow.com/a/6330109
def safe_cast(val, to_type, default=None):
	if str(val) == "":
		return 0
	try:
		return to_type(val)
	except (ValueError, TypeError):
		err("You have entered a DISGUSTING value [", '"'+ sc(val)+'"', "]", fatal=True)


def safe_num(val):
	if ("." in str(val)):
		return safe_cast(val, int)
	else:
		return safe_cast(val, float)


# Stands for "String Clean"
# Removes leading and trailing whitespace
def sc(your_string):
	return str(your_string).lstrip().rstrip()

def execute_and_save():
	global working_col_indexes_to_delete, working_row_indexes_to_delete
	for index in sorted(working_col_indexes_to_delete, reverse=True):
		del working_sheet.row[index]
	for index in sorted(working_row_indexes_to_delete, reverse=True):
		del working_sheet.column[index]
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
		if OPERATION == "REPLACE" or OPERATION == "REGREPLACE":
			# 0             1           2
			# (REG)REPLACE, search_str, replace_str
			if len(row) < 3:
				err("Not enough arguments for REPLACE operation", fatal=True)
			search_str = sc(row[1])
			replace_str = sc(row[2])
			sheet_array = working_sheet.row
			for sheet_index,row in enumerate(sheet_array):
				for location,col in enumerate(row):
					cell_value = str(row[location])
					if OPERATION == "ROW":
						working_cell_coord[1] = location
						working_cell_coord[0] = sheet_index
					else:
						working_cell_coord[1] = sheet_index
						working_cell_coord[0] = location
					working_sheet_index = sheet_index
					working_row = row
					if OPERATION == "REPLACE":
						working_row[location] = cell_value.replace(search_str, replace_str)
					else: #REGREPLACE
						working_row[location] = re.sub(search_str, replace_str, cell_value)
					sheet_array[sheet_index] = working_row

		elif OPERATION == "ROW" or OPERATION == "COL":
			# 0        1    2         3
			# ROW/COL, 1/a, operator, search
			if (len(row) < 4):
				err(OPERATION, "operation requires 3 additional arguments", fatal=True)
			sheet_array         = None
			location            = sc(row[1])
			location_array      = []        # If one or two ':' are in 'location' - to define start and end points
			search_operator     = sc(row[2])
			search_str          = sc(row[3])
			working_sheet_index = 0

			def loc_len():
				return len(location_array)

			if(":" in location):
				location_array = list(filter(None, location.split(":")))
				if loc_len() > 1:
					location = location_array[0]
				else:
					err("Your location parameter is screwed up. Ask a doctor to check out your colons :)", fatal=True)

			if OPERATION == "ROW":
				# get col array so we can easily check a row
				sheet_array = working_sheet.column
				location = int(location)-1
				if loc_len() > 1:
					location_array[1] = col2num(location_array[1])-1
				if loc_len() > 2:
					location_array[2] = col2num(location_array[2])-1
			else:
				# get row array so we can easily check a col
				sheet_array = working_sheet.row
				# must convert col letter to number
				location = col2num(location)
				if loc_len() > 1:
					location_array[1] = int(location_array[1])-1
					print(location_array)
				if loc_len() > 2:
					location_array[2] = int(location_array[2])-1

			def DeleteIfNeeded(op_condition, should_keep):
				force_true_lambda = lambda x,y: True
				if loc_len() > 1:
					should_keep = force_true_lambda if working_sheet_index < location_array[1] else should_keep
					if loc_len() > 2:
						should_keep = force_true_lambda if working_sheet_index > location_array[2] else should_keep
				if search_operator == op_condition:
					if not should_keep(search_str, cell_value):
						if OPERATION == "ROW":
							print(search_str, cell_value, should_keep(search_str, cell_value))
							working_row_indexes_to_delete.append(working_sheet_index)
						elif OPERATION == "COL":
							working_col_indexes_to_delete.append(working_sheet_index)

			for sheet_index,row_or_col in enumerate(sheet_array):
				cell_value = str(row_or_col[location])
				if OPERATION == "ROW":
					working_cell_coord[1] = location
					working_cell_coord[0] = sheet_index
				else:
					working_cell_coord[1] = sheet_index
					working_cell_coord[0] = location
				working_sheet_index = sheet_index
				# Defining the Syntax
				# x=search_str y=cell_value
				DeleteIfNeeded("is"           , lambda x,y: x == y)
				DeleteIfNeeded("!is"          , lambda x,y: x != y)
				DeleteIfNeeded("contains"     , lambda x,y: x in y)
				DeleteIfNeeded("!contains"    , lambda x,y: x not in y)
				DeleteIfNeeded("regcontains"  , lambda x,y: re.search(x, y))
				DeleteIfNeeded("!regcontains" , lambda x,y: not re.search(x, y))
				DeleteIfNeeded(">"            , lambda x,y: safe_num(y) > safe_num(x))
				DeleteIfNeeded(">="           , lambda x,y: safe_num(y) >= safe_num(x))
				DeleteIfNeeded("<"            , lambda x,y: safe_num(y) < safe_num(x))
				DeleteIfNeeded("<="           , lambda x,y: safe_num(y) <= safe_num(x))
	else:
		err("You are trying to do a", value, "operation when a sheet is not loaded.", fatal=True)

execute_and_save()

