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

import os.path

appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.9\packages\Bumblebee\extra"
if bbPath not in sys.path:
	sys.path.Add(bbPath)

import bumblebee as bb

def ReadData(ws, cellRange, byColumn):
	# get range
	if ":" in cellRange:
		origin = ws.Cells(bb.xlRange(cellRange)[1], bb.xlRange(cellRange)[0])
		extent = ws.Cells(bb.xlRange(cellRange)[3], bb.xlRange(cellRange)[2])
		rng = ws.Range[origin, extent].Value2
	else:
		# this is a named cell range
		rng = ws.Range(cellRange).Value2
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

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
byColumn = IN[3]
cellRange = IN[4]

live = False

if runMe:
	try:
		errorReport = None
		if filePath == None:
			# run excel in live mode
			if LiveStream() != None:
				xlApp = LiveStream()
				wb = xlApp.ActiveWorkbook
				live = True
			else:
				errorReport = "Open Excel file or specify path."
		else:
			# run excel from file on disk
			xlApp = SetUp(Excel.ApplicationClass())
			if os.path.isfile(str(filePath)):
				xlApp.Workbooks.open(str(filePath))
				wb = xlApp.ActiveWorkbook
			live = False				
		# get worksheet
		if sheetName == None:
			ws = xlApp.ActiveSheet
			if isinstance(cellRange, list):
				dataOut = []
				for rng in cellRange:
					dataOut.append(ReadData(ws, rng, byColumn))
			else:
				dataOut = ReadData(ws, cellRange, byColumn)
			if not live:
				ExitExcel(xlApp, wb, ws)
		elif not isinstance(sheetName, list):
			ws = xlApp.Sheets(sheetName)
			if isinstance(cellRange, list):
				dataOut = []
				for rng in cellRange:
					dataOut.append(ReadData(ws, rng, byColumn))
			else:
				dataOut = ReadData(ws, cellRange, byColumn)
			if not live:
				ExitExcel(xlApp, wb, ws)
		else:
			# process multiple sheets
			dataOut = []
			if isinstance(cellRange, list):
				for name, rng in zip(sheetName, cellRange):
					ws = xlApp.Sheets(str(name))
					dataOut.append(ReadData(ws, rng, byColumn))
			else:
				for name in sheetName:
					ws = xlApp.Sheets(str(name))
					dataOut.append(ReadData(ws, cellRange, byColumn))
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
