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

tempFilePath = IN[0]
newFilePath = IN[1]
newFileName = IN[2]
data = IN[3]
sheetName = IN[4]
tempSheetName = IN[5]
RunIt = IN[6]

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

if RunIt:
	message = None
	try:
		errorReport = None
		message = "Success!"
		
		xlApp = Excel.ApplicationClass() 
		SetUp(xlApp)
		for i in range(0, len(data), 1):
			xlApp.Workbooks.Open(unicode(tempFilePath))
			wb = xlApp.ActiveWorkbook
			ws = xlApp.Sheets(sheetName)
			
			rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(data[i]), 1))
			rng.Value = xlApp.Transpose(Array[str](data[i]))
		
			ws = xlApp.Sheets(tempSheetName)
			ws.Activate
	        
			xlApp.ActiveWorkbook.SaveAs(newFilePath + "\\" + str(newFileName) + ".xlsx")
			xlApp.ActiveWorkbook.Close(False)
			xlApp.screenUpdating = True
			Marshal.ReleaseComObject(ws)
			Marshal.ReleaseComObject(wb)
		xlApp.Quit()
		Marshal.ReleaseComObject(xlApp)
	except:
		xlApp.Quit()
		Marshal.ReleaseComObject(xlApp)
		# if error accurs anywhere in the process catch it
		import traceback
		errorReport = traceback.format_exc()
		pass
else:
	errorReport = None
	message = "Run Me is set to False."

if errorReport == None:
	OUT = OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
else:
	OUT = errorReport
