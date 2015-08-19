#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

import System
from System import Array
from System.Collections.Generic import *

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os.path
import os

appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.8\packages\Bumblebee\extra"
if bbPath not in sys.path:
	sys.path.Add(bbPath)

import bumblebee as bb

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
searchValues = IN[3]

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

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		return xlApp
	except:
		return None

def SearchValueGetRow(xlApp, ws, key):
	# get spreadhseet extents to limit search context
	originX = ws.UsedRange.Row
	originY = ws.UsedRange.Column
	boundX = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
	boundY = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
	# define search criteria
	xlAfter = ws.Cells(originY, originX)
	xlLookIn = -4163
	xlLookAt = "&H2"
	xlSearchOrder = "&H1"
	xlSearchDirection = 1
	xlMatchCase = False
	xlMatchByte = False
	xlSearchFormat = False
	# get search value and return row 
	# if value not found return None
	cell = ws.Cells.Find(key, xlAfter, xlLookIn, xlLookAt, xlSearchOrder, xlSearchDirection, xlMatchCase, xlMatchByte, xlSearchFormat)
	if cell != None:
		cellAddress = cell.Address(False, False)
		addressX = xlApp.Range(cellAddress).Row
		addressY = xlApp.Range(cellAddress).Column
		row = ws.Range[ws.Cells(addressX, originY), ws.Cells(addressX, boundY)].Value2
		return row
	else:
		return None
try:
	errorReport = None
	if runMe:
		message = None
		dataOut = []
		if filePath == None:
			# run excel in a live mode
			xlApp = LiveStream()
			wb = xlApp.ActiveWorkbook
			if sheetName == None:
				ws = xlApp.ActiveSheet
			else:
				ws = xlApp.Sheets(sheetName)
			if isinstance(searchValues, list):
				for key in searchValues:
					dataOut.append(SearchValueGetRow(xlApp, ws, key))
			else:
				dataOut = SearchValueGetRow(xlApp, ws, key)
		else:
			try:
				# open excel workbook specified at filePath
				xlApp = SetUp(Excel.ApplicationClass())
				if os.path.isfile(str(filePath)):
					xlApp.Workbooks.open(str(filePath))
					wb = xlApp.ActiveWorkbook
					ws = xlApp.Sheets(sheetName)
					if isinstance(searchValues, list):
						for key in searchValues:
							dataOut.append(SearchValueGetRow(xlApp, ws, key))
					else:
						dataOut = SearchValueGetRow(xlApp, ws, key)
					ExitExcel(filePath, xlApp, wb, ws)
			except:
				xlApp.Quit()
				Marshal.ReleaseComObject(xlApp)
				pass
	else:
		errorReport = "Set RunMe to True."
except:
		# if error accurs anywhere in the process catch it
		import traceback
		errorReport = traceback.format_exc()

#Assign your output to the OUT variable
if errorReport == None:
	OUT = dataOut
else:
	OUT = errorReport
