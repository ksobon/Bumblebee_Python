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
from collections import OrderedDict

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
		if LiveStream() != None:
			xlApp = LiveStream()
			if not isinstance(data, list):
				wb = xlApp.ActiveWorkbook
				ws = xlApp.Sheets(data.SheetName())
				WriteData(ws, data.Data(), byColumn, data.Origin())
			else:
				wb = xlApp.ActiveWorkbook
				sheetNames = list(OrderedDict.fromkeys([x.SheetName() for x in data]))
				if len(sheetNames) > wb.Sheets.Count:
					wb.Sheets.Add(After = wb.Sheets(wb.Sheets.Count), Count = len(sheetNames)-1)
				for i in range(0,len(sheetNames),1):
					wb.Worksheets[i+1].Name = sheetNames[i]
				for i in data:
					ws = xlApp.Sheets(i.SheetName())
					WriteData(ws , i.Data(), byColumn, i.Origin())
	except:
		import traceback
		errorReport = traceback.format_exc()
else:
	errorReport = None
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."
if errorReport == None:
	OUT = OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
else:
	OUT = errorReport
