#Copyright(c) 2015, David Mans, Konrad Sobon
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
		origin = ws.Cells(origin[1], origin[0])
	else:
		origin = ws.Cells(ws.UsedRange.Row, ws.UsedRange.Column)
	return origin

def GetExtent(ws, extent):
	if extent != None:
		extent = ws.Cells(extent[1], extent[0])
	else:
		extent = ws.Cells(ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)
	return extent

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

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

if runMe:
	message = None
	xlApp = SetUp(Excel.ApplicationClass())
	if os.path.isfile(str(filePath)):
		try:
			xlApp.Workbooks.open(str(filePath))
		except:
			message = "Excel might be open. Please close it!"
		if not isinstance(sheetName, list):
			wb = xlApp.ActiveWorkbook
			ws = xlApp.Sheets(sheetName)
			dataOut = ReadData(ws, GetOriginExtent(ws, origin, extent)[0], GetOriginExtent(ws, origin, extent)[1], byColumn)
			ExitExcel(filePath, xlApp, wb, ws)
		else:
			dataOut = []
			wb = xlApp.ActiveWorkbook
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
			ExitExcel(filePath, xlApp, wb, ws)
else:
	message = "Set RunMe to True."

#Assign your output to the OUT variable
if message == None:
	OUT = dataOut
else:
	OUT = message
