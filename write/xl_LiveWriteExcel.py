# Copyright(c) 2016, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys
import System
from System import Array
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os
appDataPath = os.getenv('APPDATA')
dynPath = appDataPath + r"\Dynamo\0.9"
if dynPath not in sys.path:
	sys.path.Add(dynPath)
	
bbPath = appDataPath + r"\Dynamo\0.9\packages\Bumblebee\extra"
if bbPath not in sys.path:
	try:
		sys.path.Add(bbPath)
		import bumblebee as bb
	except:
		import xml.etree.ElementTree as et
		root = et.parse(dynPath + "\DynamoSettings.xml").getroot()
		for child in root:
			if child.tag == "CustomPackageFolders":
				for path in child:
					if path not in sys.path:
						sys.path.Add(path)
		import bumblebee as bb

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

runMe = IN[0]
byColumn = IN[1]
data = IN[2]

def WriteData(ws, data, byColumn, origin):

	def FillData(x, y, x1, y1, ws, data, origin):
		if origin != None:
			x = x + origin[1]
			y = y + origin[0]
		else:
			x = x + 1
			y = y + 1
		if y1 != None:
			ws.Cells[x, y] = data[x1][y1]
		else:
			ws.Cells[x, y] = data[x1]
		return ws
	# if data is a nested list (multi column/row) use this
	if any(isinstance(item, list) for item in data):
		for i, valueX in enumerate(data):
			for j, valueY in enumerate(valueX):
				if byColumn:
					FillData(j,i,i,j, ws, data, origin)
				else:
					FillData(i,j,i,j, ws, data, origin)
	# if data is just a flat list (single column/row) use this
	else:
		for i, valueX in enumerate(data):
			if byColumn:
				FillData(i,0,i,None, ws, data, origin)
			else:
				FillData(0,i,i,None, ws, data, origin)
	return ws

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

def Flatten(*args):
    for x in args:
        if hasattr(x, '__iter__'):
            for y in Flatten(*x):
                yield y
        else:
            yield x

if isinstance(data, list):
	if any(isinstance(x, list) for x in data):
		data = list(Flatten(data))

if runMe:
	message = None
	try:
		errorReport = None
		message = "Success!"
		# check if excel is running
		if LiveStream() != None:
			xlApp = LiveStream()
			# excel is running and data is being written to single sheet
			if not isinstance(data, list):
				wb = xlApp.ActiveWorkbook
				try:
					ws = xlApp.Sheets(data.SheetName())
				# if sheet with given name doesn't exist it will be added
				except:
					ws = wb.Sheets.Add(After = wb.Sheets(wb.Sheets.Count), Count = 1)
					ws.Name = data.SheetName()
				WriteData(ws, data.Data(), byColumn, data.Origin())
			# excel is running and data is being written to multiple sheets
			else:
				wb = xlApp.ActiveWorkbook
				for i in data:
					try:
						ws = xlApp.Sheets(i.SheetName())
					# if sheet with given name doesn't exist it will be added
					except:
						ws = wb.Sheets.Add(After = wb.Sheets(wb.Sheets.Count), Count = 1)
						ws.Name = i.SheetName()
					WriteData(ws , i.Data(), byColumn, i.Origin())
		else:
			message = "No Excel session is open. Please open \nExcel to be able to use Live Excel Write node."
	except:
		# if error accurs anywhere in the process catch it
		import traceback
		errorReport = traceback.format_exc()
else:
	errorReport = None
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."
if errorReport == None:
	OUT = OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
else:
	OUT = errorReport
