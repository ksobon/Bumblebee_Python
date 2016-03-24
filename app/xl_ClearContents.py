# Copyright(c) 2016, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

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

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
clearContent = IN[3]
clearFormat = IN[4]
cellRange = IN[5]

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

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
	
	wb.SaveAs(str(filePath))
	xlApp.ActiveWorkbook.Close(False)
	xlApp.ScreenUpdating = True
	CleanUp([ws,wb,xlApp])
	return None

def ClearExcel(ws, cellRange, clearContents, clearFormats):
	# get origin and extent from range string
	# if no range supplied apply formatting to entire sheet
	if cellRange != None:
		origin = ws.Cells(bb.xlRange(cellRange)[1], bb.xlRange(cellRange)[0])
		extent = ws.Cells(bb.xlRange(cellRange)[3], bb.xlRange(cellRange)[2])
	else:
		origin = ws.Cells(ws.UsedRange.Row, ws.UsedRange.Column)
		extent = ws.Cells(ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)
	if clearContents:
		ws.Range[origin, extent].ClearContents()
	if clearFormats:
		ws.Range[origin, extent].ClearFormats()
	return ws

if runMe:
	message = None
	try:
		errorReport = None
		message = "Success!"
		if filePath == None:
			# run excel in a live mode
			xlApp = LiveStream()
			wb = xlApp.ActiveWorkbook
			if sheetName == None:
				ws = xlApp.ActiveSheet
			else:
				ws = xlApp.Sheets(sheetName)
			ClearExcel(ws, cellRange, clearContent, clearFormat)
		else:
			try:
				xlApp = SetUp(Excel.ApplicationClass())
				if os.path.isfile(str(filePath)):
					xlApp.Workbooks.open(str(filePath))
					wb = xlApp.ActiveWorkbook
					ws = xlApp.Sheets(sheetName)
					ClearExcel(ws, cellRange, clearContent, clearFormat)
					#Marshal.ReleaseComObject(extent)
					#Marshal.ReleaseComObject(origin)
					ExitExcel(filePath, xlApp, wb, ws)
				else:
					message = "Specified file doesn't exists."
			except:
				xlApp.Quit()
				Marshal.ReleaseComObject(xlApp)
	except:
		# if error accurs anywhere in the process catch it
		import traceback
		errorReport = traceback.format_exc()
		pass
else:
	errorReport = None
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."

if errorReport == None:
	OUT = OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
else:
	OUT = errorReport
