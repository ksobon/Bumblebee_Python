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

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

def ExitExcel(xlApp, wb):
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
	CleanUp([wb,xlApp])
	return None

if runMe:
	message = None
	xlApp = SetUp(Excel.ApplicationClass())
	if os.path.isfile(str(filePath)):
		try:
			xlApp.Workbooks.open(str(filePath))
		except:
			message = "Excel might be open. Please close it!"
		wb = xlApp.ActiveWorkbook
		dataOut = []
		for i in range(0, xlApp.Sheets.Count, 1):
			dataOut.append(xlApp.Sheets(i+1).Name)
	ExitExcel(xlApp, wb)	
else:
	message = "Set RunMe to True."

#Assign your output to the OUT variable
if message == None:
	OUT = dataOut
else:
	OUT = message
