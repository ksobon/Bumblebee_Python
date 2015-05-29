#Copyright(c) 2015, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

import os
import os.path
appDataPath = os.getenv('APPDATA')
bbPath = appDataPath + r"\Dynamo\0.8\packages\Bumblebee\extra"
if bbPath not in sys.path:
	sys.path.Add(bbPath)

import System
from System import Array
from System.Collections.Generic import *

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
from System.Runtime.InteropServices import Marshal

import bumblebee as bb
import string
import re

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

filePath = IN[0]
runMe = IN[1]
sheetName = IN[2]
cellRange = IN[3]
formatConditions = IN[4]

def SetUp(xlApp):
	# supress updates and warning pop ups
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	xlApp.ScreenUpdating = False
	return xlApp

def CleanUp(_list):
	# clean up before exiting excel, if any COM object remains
	# unreleased then excel crashes on open following time
	if isinstance(_list, list):
		for i in _list:
			try:
				Marshal.ReleaseComObject(i)
			except:
				pass
	else:
		try:
			Marshal.ReleaseComObject(_list)
		except:
			pass
	return None

def ExitExcel(filePath, xlApp, wb, ws):
	# save any changes and clean up excel sessions
	wb.SaveAs(str(filePath))
	xlApp.ActiveWorkbook.Close(False)
	xlApp.ScreenUpdating = True
	CleanUp([ws,wb,xlApp])
	return None

def ConditionFormatCells(origin, extent, ws, formatConditions):
	
	def AddFormatCondition(origin=None, extent=None, ws=None, formatConditions=None, index=None):
		fcType = formatConditions.FormatConditionType()
		if fcType == 1:
			operatorType = formatConditions.OperatorType()
			values = formatConditions.Values()
			if operatorType == 2 or operatorType == 1:
				ws.Range[origin, extent].FormatConditions.Add(fcType, operatorType, values[1], values[0])
			else:
				ws.Range[origin, extent].FormatConditions.Add(fcType, operatorType, values)
				
		if fcType == 2:
			operatorType = formatConditions.OperatorType()
			expression = formatConditions.Expression()
			ws.Range[origin, extent].FormatConditions.Add(fcType, operatorType, expression)
			
		if fcType == "2Color":
			ws.Range[origin, extent].FormatConditions.AddColorScale(ColorScaleType = 2)
		
		if fcType == "3Color":
			ws.Range[origin, extent].FormatConditions.AddColorScale(ColorScaleType = 3)
		
		if fcType == "TopPercentile":
			ws.Range[origin, extent].FormatConditions.AddTop10()
			if index == None:
				index = 1
			else:
				index = index + 1
			ws.Range[origin, extent].FormatConditions(index).Percent = formatConditions.Percent()
			ws.Range[origin, extent].FormatConditions(index).Rank = formatConditions.Rank()
			ws.Range[origin, extent].FormatConditions(index).TopBottom = formatConditions.TopBottom()

		return ws
		
	def FormatGraphics(origin=None, extent=None, ws=None, formatConditions=None, index=None):
		if index == None:
			index = 1
		else:
			index = index + 1
			
		if formatConditions.FormatConditionType() == "2Color":
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(1).Type = formatConditions.MinType()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(1).FormatColor.Color = formatConditions.MinColor()
			if formatConditions.MinType() != 1:
				ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(1).Value = formatConditions.MinValue()
			if formatConditions.MaxType() != 2:
				ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(2).Value = formatConditions.MaxValue()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(2).Type = formatConditions.MaxType()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(2).FormatColor.Color = formatConditions.MaxColor()
			
		elif formatConditions.FormatConditionType() == "3Color":
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(1).Type = formatConditions.MinType()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(1).FormatColor.Color = formatConditions.MinColor()
			if formatConditions.MinType() != 1:
				ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(1).Value = formatConditions.MinValue()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(2).Type = formatConditions.MidType()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(2).FormatColor.Color = formatConditions.MidColor()
			ws.Range[origin, extent].FormatConditions(index).ColorScaleCriteria(2).Value = formatConditions.MidValue()

		else:
			fillStyle = formatConditions.GraphicStyle().fillStyle
			textStyle = formatConditions.GraphicStyle().textStyle
			borderStyle = formatConditions.GraphicStyle().borderStyle
			
			if fillStyle.backgroundColor != None:
				ws.Range[origin, extent].FormatConditions(index).Interior.Color = fillStyle.BackgroundColor()
			if fillStyle.patternType != None:
				ws.Range[origin, extent].FormatConditions(index).Interior.Pattern = fillStyle.PatternType()
			if fillStyle.patternColor != None:
				ws.Range[origin, extent].FormatConditions(index).Interior.PatternColor = fillStyle.PatternColor()
			ws.Range[origin, extent].FormatConditions(index).StopIfTrue = False
		return ws

	ws.Range[origin, extent].FormatConditions.Delete()
	if not isinstance(formatConditions, list):
		AddFormatCondition(origin, extent, ws, formatConditions)
		FormatGraphics(origin, extent, ws, formatConditions, None)
	else:
		for index, value in enumerate(formatConditions):
			AddFormatCondition(origin, extent, ws, value)
			FormatGraphics(origin, extent, ws, value, index)
	return ws
	
if runMe:
	message = None
	xlApp = SetUp(Excel.ApplicationClass())
	try:
		if os.path.isfile(str(filePath)):
			xlApp.Workbooks.open(str(filePath))
			wb = xlApp.ActiveWorkbook
			ws = xlApp.Sheets(sheetName)
			if not isinstance(cellRange, list):
				origin = ws.Cells(bb.xlRange(cellRange)[1], bb.xlRange(cellRange)[0])
				extent = ws.Cells(bb.xlRange(cellRange)[3], bb.xlRange(cellRange)[2])
				ConditionFormatCells(origin, extent, ws, formatConditions)
				Marshal.ReleaseComObject(extent)
				Marshal.ReleaseComObject(origin)
			else:
				for index, (range, format) in enumerate(zip(cellRange, formatConditions)):
					origin = ws.Cells(bb.xlRange(range)[1], bb.xlRange(range)[0])
					extent = ws.Cells(bb.xlRange(range)[3], bb.xlRange(range)[2])
					ConditionFormatCells(origin, extent, ws, format)
					Marshal.ReleaseComObject(extent)
					Marshal.ReleaseComObject(origin)
			ExitExcel(filePath, xlApp, wb, ws)
		else:
			message = "No file exists. Please Use Write Data Node to create file first before formatting it."
	except:
		xlApp.Quit()
		Marshal.ReleaseComObject(xlApp)
		message = "Something went wrong. Please check \n your inputs and try again."
		pass
else:
	message = "Run Me is set to False. Please set \nto True if you wish to write data \nto Excel."

#Assign your output to the OUT variable
if message == None:
	OUT = "Success!"
else:
	OUT = '\n'.join('{:^35}'.format(s) for s in message.split('\n'))
