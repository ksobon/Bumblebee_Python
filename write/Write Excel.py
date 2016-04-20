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
					path_bb = path.text + "\Bumblebee\extra"
					if path not in sys.path:
						sys.path.Add(path_bb)
		import bumblebee as bb

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
byColumn = IN[2]
data = IN[3]

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

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

def ExitExcel(filePath, xlApp, wb, ws):
	# clean up before exiting excel, if any COM object remains
	# unreleased then excel crashes on open following time
	def CleanUp(_list):
		if isinstance(_list, list):
			for i in _list:
				Marshal.ReleaseComObject(i)
		else:
			Marshal.ReleaseComObject(_list)
		return None
	
	wb.SaveAs(unicode(filePath))
	xlApp.ActiveWorkbook.Close(False)
	xlApp.ScreenUpdating = True
	CleanUp([ws,wb,xlApp])
	return None

def Flatten(*args):
    for x in args:
        if hasattr(x, '__iter__'):
            for y in Flatten(*x):
                yield y
        else:
            yield x

def WorksheetExists(wb, name):
	for i in wb.Sheets:
		if i.Name == name:
			return True
			break
		else:
			continue
	return False

if isinstance(data, list):
	if any(isinstance(x, list) for x in data):
		data = list(Flatten(data))

live = False

if runMe:
	try:
		errorReport = None
		if filePath == None:
			# run excel in live mode
			xlApp = LiveStream()
			live = True
			wb = xlApp.ActiveWorkbook
		else:
			# run excel from a file on disk
			xlApp = SetUp(Excel.ApplicationClass())
			live = False
			# if file exists open it
			if os.path.isfile(unicode(filePath)):
				xlApp.Workbooks.open(unicode(filePath))
				wb = xlApp.ActiveWorkbook
			# if file doesn't exist just make a new one
			else:
				wb = xlApp.Workbooks.Add()
				if not isinstance(data, list):
					# add and rename worksheet
					ws = wb.Worksheets[1]
					ws.Name = data.SheetName()
				else:
					for i in data:
						# if worksheet doesn't exist add it and name it
						if not WorksheetExists(wb, i.SheetName()):
							wb.Sheets.Add(After = wb.Sheets(wb.Sheets.Count), Count = 1)
							ws = wb.Worksheets[wb.Sheets.Count]
							ws.Name = i.SheetName()
		# data is a flat list - single sheet gets written
		if not isinstance(data, list):
			if WorksheetExists(wb, data.SheetName()):
				ws = xlApp.Sheets(data.SheetName())
			else:
				wb.Sheets.Add(After = wb.Sheets(wb.Sheets.Count), Count = 1)
				ws = wb.Worksheets[wb.Sheets.Count]
				ws.Name = data.SheetName()
			WriteData(ws, data.Data(), byColumn, data.Origin())
			if not live:
				ExitExcel(filePath, xlApp, wb, ws)
		# data is a nested list - multiple sheets are written
		else:
			sheetNameSet = set([x.SheetName() for x in data])
			for i in data:
				if WorksheetExists(wb, i.SheetName()):
					ws = xlApp.Sheets(i.SheetName())
				else:
					wb.Sheets.Add(After = wb.Sheets(wb.Sheets.Count), Count = 1)
					ws = wb.Worksheets[wb.Sheets.Count]
					ws.Name = i.SheetName()
				WriteData(ws, i.Data(), byColumn, i.Origin())
			if not live:
				ExitExcel(filePath, xlApp, wb, ws)
				
	except:
		xlApp.Quit()
		Marshal.ReleaseComObject(xlApp)
		# if error accurs anywhere in the process catch it
		import traceback
		errorReport = traceback.format_exc()
		pass
else:
	errorReport = "Set RunMe to True."

#Assign your output to the OUT variable
if errorReport == None:
	OUT = "Success!"
else:
	OUT = errorReport
