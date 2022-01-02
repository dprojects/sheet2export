# -*- coding: utf-8 -*-

# FreeCAD macro for spreadsheet export
# Author: Darek L (aka dprojects)
# Version: 2022.01.02
# Latest version: https://github.com/dprojects/sheet2export

import FreeCAD, Draft, Spreadsheet

# ###################################################################################################################
# Main Settings ( CHANGE HERE IF NEEDED )
# ###################################################################################################################


# File type:
# "csv" - Comma-separated values (.csv file)
# "html" - HyperText Markup Language (.html file)
# "json" - JavaScript Object Notation (.json file), see e.g. json2table.com
# "md" - MarkDown (.md file), see e.g. dillinger.io
sFileType = "html"

# Export type:
# "a" - all spreadsheet objects
# "s" - selected spreadsheet only
sExportType = "a"

# File path:
# "~" - user home folder
# "./" - current macro folder
# or set Your custom path with write permissions
sFilePath = "./"


# ###################################################################################################################
# Additional Settings ( CHANGE HERE IF NEEDED )
# ###################################################################################################################

# Empty cell content:
sEmptyCell = "&nbsp;" # "&nbsp;" but this can be any character "-", "N/A" or "0"

# Separator for CSV:
sSepCSV = ","

# CSS table border decoration:
sBorderSize = "1px" # set to "0" to turn it off
sBorderType = "dotted" # dotted dashed solid
sBorderPosition = "bottom" # bottom left right top
sBorderColor = "#000000"


# ###################################################################################################################
# Autoconfig - define globals ( NOT CHANGE HERE )
# ###################################################################################################################


# set reference point to Active Document
gAD = FreeCAD.activeDocument()

# get all objects from 3D model
gOBs = gAD.Objects

# init output file name
gFile = "result" # will be overwritten later

# init spreadsheet object
gSheet = gAD # will be overwritten later

# init output result
gOUT = ""

# console print separator
gSepC = "\n ================================================================ \n"


# ###################################################################################################################
# Databases
# ###################################################################################################################


# cell properties
dbCPC = dict() # content
dbCPA = dict() # alignment
dbCPS = dict() # style
dbCPB = dict() # background
dbCPRS = dict() # row span
dbCPCS = dict() # column span

# max
dbMaxR = 0 # row
dbMaxC = 0 # column

# spreadsheet key
dbSKV = dict() # value for letters
dbSKL = dict() # letters for value


# ###################################################################################################################
# Support for errors
# ###################################################################################################################


# ###################################################################################################################
def showError(iObj, iPlace, iError):

	FreeCAD.Console.PrintMessage(gSepC)
	
	try:
		FreeCAD.Console.PrintMessage("ERROR: ")
		FreeCAD.Console.PrintMessage(" | ")
		FreeCAD.Console.PrintMessage(str(iObj.Label))
		FreeCAD.Console.PrintMessage(" | ")
		FreeCAD.Console.PrintMessage(str(iPlace))
		FreeCAD.Console.PrintMessage(" | ")
		FreeCAD.Console.PrintMessage(str(iError))
		
	except:
		FreeCAD.Console.PrintMessage("FATAL ERROR, or even worse :-)")
		
	FreeCAD.Console.PrintMessage(gSepC)
	
	return 0


# ###################################################################################################################
# CSV file format ( COPY AND CHANGE TO ADD NEW FILE FORMAT )
# ###################################################################################################################


# ###################################################################################################################
def CSVbegin():
	global gOUT

	gOUT += ''


# ###################################################################################################################
def CSVend():
	global gOUT

	gOUT += ''


# ###################################################################################################################
def CSVrowOpen():
	global gOUT

	gOUT += ''


# ###################################################################################################################
def CSVrowClose():
	global gOUT

	gOUT += '\n'


# ###################################################################################################################
def CSVempty(iKey, iC, iR):
	global gOUT

	gOUT += sSepCSV


# ###################################################################################################################
def CSVcell(iKey, iCell, iC, iR):
	global gOUT

	gOUT += str(iCell)
	gOUT += sSepCSV


# ###################################################################################################################
# HTML file format
# ###################################################################################################################


# ###################################################################################################################
def HTMLbegin():
	global gOUT

	# there is no need to add html document header here because if the file is html table 
	# only the file is correctly parsed by browser, moreover this is easier to copy the 
	# file content and place it to the post or other web page
	gOUT += '<TABLE>'


# ###################################################################################################################
def HTMLend():
	global gOUT

	gOUT += "</TABLE>"


# ###################################################################################################################
def HTMLrowOpen():
	global gOUT

	gOUT += "<TR>"


# ###################################################################################################################
def HTMLrowClose():
	global gOUT

	gOUT += "</TR>"


# ###################################################################################################################
def HTMLempty(iKey, iC, iR):
	global gOUT

	gOUT += '<TD '
	gOUT += 'style="border-' 
	gOUT += str(sBorderPosition) + ':'
	gOUT += str(sBorderSize) + ' '
	gOUT += str(sBorderType) + ' '
	gOUT += str(sBorderColor) + ';"'
	gOUT += '>'
	gOUT += str(sEmptyCell)
	gOUT += "</TD>"


# ###################################################################################################################
def HTMLcell(iKey, iCell, iC, iR):
	global gOUT

	gOUT += '<TD '

	try:
		gOUT += 'colspan="'+str(dbCPCS[iKey]) + '" '
	except:
		gOUT += ''

	try:
		gOUT += 'rowspan="'+str(dbCPRS[iKey]) + '" '
	except:
		gOUT += ''

	gOUT += 'style="'

	gOUT += 'border-' 
	gOUT += str(sBorderPosition) + ':'
	gOUT += str(sBorderSize) + ' '
	gOUT += str(sBorderType) + ' '
	gOUT += str(sBorderColor) + ';'
	
	try:
		gOUT += 'text-align:'+str(dbCPA[iKey]).split("|")[0] + ';'
	except:
		gOUT += ''
	
	try:
		gOUT += 'background-color:'+str(dbCPB[iKey]) + ';'
	except:
		gOUT += ''
	
	try:
		gOUT += 'font-weight:'+str(dbCPS[iKey]) + ';'
	except:
		gOUT += ''
	

	gOUT += '"'
	gOUT += '>'
	gOUT += str(iCell)
	gOUT += "</TD>"


# ###################################################################################################################
# JSON file format
# ###################################################################################################################


# ###################################################################################################################
def JSONbegin():
	global gOUT

	gOUT += '['


# ###################################################################################################################
def JSONend():
	global gOUT

	gOUT = gOUT[:-1]
	gOUT += ']'


# ###################################################################################################################
def JSONrowOpen():
	global gOUT

	gOUT += '{'


# ###################################################################################################################
def JSONrowClose():
	global gOUT

	gOUT = gOUT[:-1]
	gOUT += '},'


# ###################################################################################################################
def JSONempty(iKey, iC, iR):
	global gOUT

	key = str(dbSKL[str(iC)])
	gOUT += '"' + str(key) + '":'
	gOUT += '""'
	gOUT += ','


# ###################################################################################################################
def JSONcell(iKey, iCell, iC, iR):
	global gOUT

	key = str(dbSKL[str(iC)])
	gOUT += '"' + str(key) + '":'
	gOUT += '"' + str(iCell) + '"'
	gOUT += ','


# ###################################################################################################################
# MarkDown file format
# ###################################################################################################################


# ###################################################################################################################
def MDbegin():
	global gOUT

	c = 1
	while c < dbMaxC + 1:
		gOUT += '|   '
		c = c + 1

	gOUT += '|\n'

	c = 1
	while c < dbMaxC + 1:

		# set alignment
		try:	

			key = str(dbSKL[str(c)]) + str(1)
			a = str(dbCPA[key]).split("|")[0]

			if a == "left":
				gOUT += '|:--'
			if a == "right":
				gOUT += '|--:'
			if a == "center":
				gOUT += '|:-:'
		except:
			gOUT += '|---'

		c = c + 1

	gOUT += '|\n'


# ###################################################################################################################
def MDend():
	global gOUT

	gOUT += ''


# ###################################################################################################################
def MDrowOpen():
	global gOUT

	gOUT += ''


# ###################################################################################################################
def MDrowClose():
	global gOUT

	gOUT += '   |'
	gOUT += '\n'


# ###################################################################################################################
def MDempty(iKey, iC, iR):
	global gOUT

	gOUT += '|   '


# ###################################################################################################################
def MDcell(iKey, iCell, iC, iR):
	global gOUT

	gOUT += '|   '
	gOUT += str(iCell)


# ###################################################################################################################
# Database write conroller
# ###################################################################################################################


# ###################################################################################################################
def getKey(iC, iR):

	letters = "ZABCDEFGHIJKLMNOPQRSTUVWXYZ" # max 26

	mod = int((iC % 26))
	div = int((iC - 1) / 26) 

	keyC = ""

	if iC < 27:
		keyC = letters[iC]
	else:
		keyC = letters[div] + letters[mod]

	keyR = str(iR)

	key = keyC + keyR

	# for given column and row it returns
	# spreadsheet key for cell like e.g. A5, AG125 etc
	return str(key)


# ###################################################################################################################
def setSK():

	# set spreadsheet keys databases
	c = 1
	while c < 703:

		# get key for first row and remove number
		key = getKey(c, 1)[:-1]
		
		# set spreadsheet key value
		dbSKV[key] = c

		# set spreadsheet key letter
		dbSKL[str(c)] = key

		# go to next column
		c = c + 1


# ###################################################################################################################
def setDB():

	# import regular expressions
	import re

	# refer to globals
	global dbMaxR
	global dbMaxC
	
	# XML parse part from python doc
	import xml.etree.ElementTree as ET
	result = str(gSheet.cells.Content)
	root = ET.fromstring(result)

	# set only available data
	for child in root:
		root2 = dict(child.attrib)
		
		# skip data not related to cells
		try:
			key = root2["address"]
		except:
			continue

		try:
			# the XML parse may not be consistent with the FreeCAD spreadsheet objects,
			# the XML may contains extra characters like "=" or '' so you have to write 
			# the FreeCAD content not the XML content with the extra characters
			dbCPC[key] = gSheet.get(key)
		except:
			skip = 1

		try:
			dbCPA[key] = root2["alignment"]
		except:
			skip = 1

		try:
			dbCPS[key] = root2["style"]
		except:
			skip = 1

		try:
			dbCPB[key] = root2["backgroundColor"]
		except:
			skip = 1

		try:
			dbCPRS[key] = root2["rowSpan"]
		except:
			skip = 1

		try:
			dbCPCS[key] = root2["colSpan"]
		except:
			skip = 1

		# width is not set because web pages and other formats has its own 
		# page size, for advance science data the spreadsheet can be even 
		# for ZZ column, so its not make any sense to recalculate it, 
		# width of column should be in auto mode, as small as possible 
		# but keep the text readable and possible to print, 
		# columns can be adjusted manually if needed

	# set max row and max column
	# search keys in content db
	for k in dbCPC.keys():
		
		# split spreadsheet key to word and number
		res = [re.findall(r'(\w+?)(\d+)', str(k))[0] ]

		w = str(res[0][0]) # word column
		n = int(res[0][1]) # number rows
		
		if int(dbSKV[w]) > int(dbMaxC):
			dbMaxC = int(dbSKV[w])
		
		if int(n) > int(dbMaxR):
			dbMaxR = int(n)

	# search keys also in background db
	# this can be page separator line using background color
	for k in dbCPB.keys():
		
		# split spreadsheet key to word and number
		res = [re.findall(r'(\w+?)(\d+)', str(k))[0] ]

		w = str(res[0][0]) # word column
		n = int(res[0][1]) # number rows
		
		if int(dbSKV[w]) > int(dbMaxC):
			dbMaxC = int(dbSKV[w])
		
		if int(n) > int(dbMaxR):
			dbMaxR = int(n)
	

# ###################################################################################################################
def resetDB():

	# reset db
	global dbCPC
	global dbCPA
	global dbCPS
	global dbCPB
	global dbCPRS
	global dbCPCS

	dbCPC.clear() # content
	dbCPA.clear() # alignment
	dbCPS.clear() # style
	dbCPB.clear() # background
	dbCPRS.clear() # row span
	dbCPCS.clear() # column span

	# reset output
	global gOUT

	gOUT = ""

	# max
	global dbMaxR
	global dbMaxC

	dbMaxR = 0 # row
	dbMaxC = 0 # column


# ###################################################################################################################
# File format selector
# ###################################################################################################################


# ###################################################################################################################
def selectBegin():

	if sFileType == "csv":
		CSVbegin()

	if sFileType == "html":
		HTMLbegin()

	if sFileType == "json":
		JSONbegin()

	if sFileType == "md":
		MDbegin()


# ###################################################################################################################
def selectEnd():

	if sFileType == "csv":
		CSVend()

	if sFileType == "html":
		HTMLend()

	if sFileType == "json":
		JSONend()

	if sFileType == "md":
		MDend()


# ###################################################################################################################
def selectRowOpen():

	if sFileType == "csv":
		CSVrowOpen()

	if sFileType == "html":
		HTMLrowOpen()

	if sFileType == "json":
		JSONrowOpen()

	if sFileType == "md":
		MDrowOpen()


# ###################################################################################################################
def selectRowClose():

	if sFileType == "csv":
		CSVrowClose()

	if sFileType == "html":
		HTMLrowClose()

	if sFileType == "json":
		JSONrowClose()

	if sFileType == "md":
		MDrowClose()


# ###################################################################################################################
def selectEmpty(iKey, iC, iR):

	if sFileType == "csv":
		CSVempty(iKey, iC, iR)

	if sFileType == "html":
		HTMLempty(iKey, iC, iR)

	if sFileType == "json":
		JSONempty(iKey, iC, iR)

	if sFileType == "md":
		MDempty(iKey, iC, iR)


# ###################################################################################################################
def selectCell(iKey, iCell, iC, iR):

	if sFileType == "csv":
		CSVcell(iKey, iCell, iC, iR)

	if sFileType == "html":
		HTMLcell(iKey, iCell, iC, iR)

	if sFileType == "json":
		JSONcell(iKey, iCell, iC, iR)

	if sFileType == "md":
		MDcell(iKey, iCell, iC, iR)


# ###################################################################################################################
# Set output
# ###################################################################################################################


# ###################################################################################################################
def setOUTPUT():

	# set begin of the spreadsheet table
	selectBegin()
	
	# set variables for loop
	colSpan = 0
	rowSpan = 0
	vCell = ""
	c = 1
	r = 1
	
	# just simple walk thru the matrix and setting each cell
	while r <= dbMaxR:

		FreeCAD.Console.PrintMessage(".")
		FreeCAD.Console.PrintMessage("")

		# set row extra properties
		selectRowOpen()
		
		# go thru columns for given row
		while c <= dbMaxC:

			# get access point
			vKey = str(dbSKL[str(c)]) + str(r)

			try:
				
				# get content
				vCell = str(dbCPC[vKey])

				# set colspan before you set the cell
				try:
					colSpan = int(dbCPCS[vKey])
					rowSpan = int(dbCPRS[vKey])
				except:
					skip = 1

				# set the cell content
				selectCell(vKey, vCell, c, r)

			except:

				# if there was no such cell access point it will be empty cell
				# if there is open colspan this should be skipped
				if colSpan == 0 and rowSpan == 0:
					selectEmpty(vKey, c, r)

				# colspan is not supported by the file type
				if sFileType != "html":
					if colSpan != 0 or  rowSpan != 0:
						selectEmpty(vKey, c, r)

			# if the cell was written and there is colspan open
			if colSpan > 0:
				colSpan = colSpan - 1 
		
			# just go to next column
			c = c + 1

		# add extra close row properties
		selectRowClose()
		
		if rowSpan > 0:
			rowSpan = rowSpan - 1 		

		# set variables for next row
		c = 1
		r = r + 1

	# set end of the spreadsheet table
	selectEnd()

	# set info
	FreeCAD.Console.PrintMessage("done.")


# ###################################################################################################################
# Save spreadsheet data to file
# ###################################################################################################################


# ###################################################################################################################
def saveToDisk():

	import os, sys
	from os.path import expanduser
	
	vRoot = expanduser(sFilePath)
	vFileName = str(gFile) + "." + str(sFileType)
	vFile = os.path.join(vRoot, vFileName)
	
	FreeCAD.Console.PrintMessage("\n")
	FreeCAD.Console.PrintMessage("Saved at: ")
	FreeCAD.Console.PrintMessage(vFile)
	FreeCAD.Console.PrintMessage("\n")
	
	with open(vFile, 'w') as vFH:
		vFH.write("%s" % gOUT)


# ###################################################################################################################
# MAIN TASKS
# ###################################################################################################################

# ###################################################################################################################
def runTasks():

	try:
		setDB()
	except:
		showError(gSheet, "setDB" , "Databese is not set correctly.")
		
	try:
		setOUTPUT()
	except:
		showError(gSheet, "setOUTPUT" , "Output is not set correctly.")
		
	try:	
		saveToDisk()
	except:
		showError(gSheet, "saveToDisk" , "File is not exported correctly.")


# ###################################################################################################################
# MAIN
# ###################################################################################################################

# set info separator
FreeCAD.Console.PrintMessage(gSepC)

# set spreadsheet key databases
try:
	setSK()
except:
	showError(gAD, "setSK" , "Spreadsheet key databases is not set correctly.")

# for selected
if sExportType == "s":
	try:
		# try set selected spreadsheet
		gSheet = FreeCADGui.Selection.getSelection()[0]
	except:
		showError(gAD, "main", "Please select spreadsheet to export.")

		# check if this is correct spreadsheet object
		if gSheet.isDerivedFrom("Spreadsheet::Sheet"):

			# set output filename
			gFile = gAD.Label + " - " + gSheet.Label

			# set info
			FreeCAD.Console.PrintMessage("\n")
			FreeCAD.Console.PrintMessage("Exporting: ")
			FreeCAD.Console.PrintMessage(gSheet.Label + " ")

			# create output file
			runTasks()

# for all spreadsheets
elif sExportType == "a":

	# search all objects and export spreadsheets
	for obj in gOBs:

		# try set spreadsheet
		gSheet = obj

		# check if this is correct spreadsheet object
		if gSheet.isDerivedFrom("Spreadsheet::Sheet"):

				# set output filename
			gFile = gAD.Label + " - " + gSheet.Label
		else:
			continue

		# set info
		FreeCAD.Console.PrintMessage("\n")
		FreeCAD.Console.PrintMessage("Exporting: ")
		FreeCAD.Console.PrintMessage(gSheet.Label + " ")
		
		# create output file
		resetDB()
		runTasks()

else:
	showError(gAD, "main", "Please set sExportType correctly.")


# set info separator
FreeCAD.Console.PrintMessage(gSepC)


# ###################################################################################################################
# Reload to see changes
# ###################################################################################################################


gAD.recompute()

