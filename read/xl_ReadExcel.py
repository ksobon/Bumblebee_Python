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

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
byColumn = IN[3]
origin = IN[4]
extent = IN[5]

def ReadData(ws, origin, extent, byColumn):

	rng = ws.Range[origin, extent].Value2
	if not byColumn:
		dataOut = [[] for i in range(rng.GetUpperBound(0))]
		for i in range(rng.GetLowerBound(0)-1, rng.GetUpperBound(0), 1):
			for j in range(rng.GetLowerBound(1)-1, rng.GetUpperBound(1), 1):
				dataOut[i].append(rng[i,j])
		return dataOut
	else:
		dataOut = [[] for i in range(rng.GetUpperBound(1))]
		for i in range(rng.GetLowerBound(1)-1, rng.GetUpperBound(1), 1):
			for j in range(rng.GetLowerBound(0)-1, rng.GetUpperBound(0), 1):
				dataOut[i].append(rng[j,i])
		return dataOut

def GetOrigin(ws, origin):
	if origin != None:
		origin = ws.Cells(bb.CellIndex(origin)[1], bb.CellIndex(origin)[0])
	else:
		origin = ws.Cells(ws.UsedRange.Row, ws.UsedRange.Column)
	return origin

def GetExtent(ws, extent):
	if extent != None:
		extent = ws.Cells(bb.CellIndex(extent)[1], bb.CellIndex(extent)[0])
	else:
		extent = ws.Cells(ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)
	return extent

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

def ExitExcel(xlApp, wb, ws):
	# clean up before exiting excel, if any COM object remains
	# unreleased then excel crashes on open following time
	def CleanUp(_list):
		if isinstance(_list, list):
			for i in _list:
				Marshal.ReleaseComObject(i)
		else:
			Marshal.ReleaseComObject(_list)
		return None
		
	xlApp.ActiveWorkbook.Close(False)
	xlApp.ScreenUpdating = True
	CleanUp([ws,wb,xlApp])
	return None

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

live = False

if runMe:
	try:
		errorReport = None
		message = None
		if filePath == None:
			# run excel in live mode
			xlApp = LiveStream()
			live = True
		else:
			# run excel from file on disk
			xlApp = SetUp(Excel.ApplicationClass())
			if os.path.isfile(unicode(filePath)):
				xlApp.Workbooks.open(unicode(filePath))
			live = False				
		# get workbook
		wb = xlApp.ActiveWorkbook
		# get worksheet
		if sheetName == None:
			ws = xlApp.ActiveSheet
			dataOut = ReadData(ws, GetOrigin(ws, origin), GetExtent(ws, extent), byColumn)
			if not live:
				ExitExcel(xlApp, wb, ws)
		elif not isinstance(sheetName, list):
			ws = xlApp.Sheets(sheetName)
			dataOut = ReadData(ws, GetOrigin(ws, origin), GetExtent(ws, extent), byColumn)
			if not live:
				ExitExcel(xlApp, wb, ws)
		else:
			# process multiple sheets
			dataOut = []
			if isinstance(origin, list):
				if isinstance(extent, list):
					for index, (name, oValue, eValue) in enumerate(zip(sheetName, origin, extent)):
						ws = xlApp.Sheets(str(name))
						dataOut.append(ReadData(ws, GetOrigin(ws, oValue), GetExtent(ws, eValue), byColumn))
				else:
					for index, (name, oValue) in enumerate(zip(sheetName, origin)):
						ws = xlApp.Sheets(str(name))
						dataOut.append(ReadData(ws, GetOrigin(ws, oValue), GetExtent(ws, extent), byColumn))
			else:
				if isinstance(extent, list):
					for index, (name, eValue) in enumerate(zip(sheetName, extent)):
						ws = xlApp.Sheets(str(name))
						dataOut.append(ReadData(ws, GetOrigin(ws, origin), GetExtent(ws, eValue), byColumn))
				else:
					for index, name in enumerate(sheetName):
						ws = xlApp.Sheets(str(name))
						dataOut.append(ReadData(ws, GetOrigin(ws, origin), GetExtent(ws, extent), byColumn))
			if not live:
				ExitExcel(xlApp, wb, ws)	
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
	OUT = dataOut
else:
	OUT = errorReport
