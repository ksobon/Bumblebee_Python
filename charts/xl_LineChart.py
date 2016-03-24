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
import string
import re

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

runMe = IN[0]
sheetName = IN[1]
size = IN[2]
title = IN[3]
dataRange = IN[4]
chartType = IN[5]
legendStyle = IN[6]
chartStyle = IN[7]
graphStyle = IN[8]

def LiveStream():
	try:
		xlApp = Marshal.GetActiveObject("Excel.Application")
		xlApp.Visible = True
		xlApp.DisplayAlerts = False
		xlApp.ScreenUpdating = False
		return xlApp
	except:
		return None

def GetWidthHeight(origin, extent, ws):
	left = origin.Left
	top = origin.Top
	width = ws.Range[origin, extent].Width
	height = ws.Range[origin, extent].Height
	return [left, top, width, height]

def FormatBorder(border, bStyle):
	if bStyle.lineType != None:
		border.LineStyle = bStyle.LineType()
	if bStyle.weight != None:
		border.Weight = bStyle.Weight()
	if bStyle.color != None:
		border.Color = bStyle.Color()
	return None

def FormatText(text, tStyle):
	if tStyle.name != None:
		text.Name = tStyle.Name()
	if tStyle.size != None:
		text.Size = tStyle.Size()
	if tStyle.color != None:
		text.Color = tStyle.Color()
	if tStyle.bold != None:
		text.Bold = tStyle.Bold()
	if tStyle.italic != None:
		text.Italic = tStyle.Italic()
	if tStyle.underline != None:
		text.Underline = tStyle.Underline()
	if tStyle.strikethrough != None:
		text.Strikethrough = tStyle.Strikethrough()
	return None

def FormatFill(fill, fStyle):
	if fStyle.patternType != None:
		fill.Pattern = fStyle.PatternType()	
	if fStyle.patternColor != None:
		fill.PatternColor = fStyle.PatternColor()	
	if fStyle.backgroundColor != None:
		fill.Color = fStyle.BackgroundColor()
	return None

def ApplyGraphStyle(chartType, series, graphStyle):
	# format label style
	if graphStyle.labelStyle != None:
		if series.HasDataLabels == True:
			series.DataLabels().Delete()
			series.HasDataLabels = True
		else:
			series.HasDataLabels = True
		gls = graphStyle.labelStyle
		
		if gls.seriesName != None:
			series.DataLabels().ShowSeriesName = gls.SeriesName()
		if gls.value != None:
			series.DataLabels().ShowValue = gls.Value()
		if gls.legendKey != None:
			series.DataLabels().ShowLegendKey = gls.LegendKey()
		if gls.separator != None:
			series.DataLabels().Separator = gls.Separator()
		if gls.labelPosition != None:
			series.DataLabels().Position = gls.LabelPosition()
	
		# set label fill settings
		if gls.fillStyle != None:
			FormatFill(series.DataLabels().Interior, gls.fillStyle)
		# set label text style
		if gls.textStyle != None:
			FormatText(series.DataLabels().Font, gls.textStyle)
		# set label border style
		if gls.borderStyle != None:
			FormatBorder(series.DataLabels().Border, gls.borderStyle)
	else:
		series.HasDataLabels = False
		
	# format line style
	if graphStyle.lineStyle != None:
		ls = graphStyle.lineStyle
		series.Format.Line.Visible = True
		if ls.color != None:
			series.Format.Line.ForeColor.RGB = ls.Color()
		if ls.weight != None:
			series.Format.Line.Weight = ls.Weight()
		if ls.compoundLineType != None:
			series.Format.Line.Style = ls.CompoundLineType()
		if ls.lineType != None:
			series.Format.Line.DashStyle = ls.LineType()
		if ls.smooth != None:
			series.Smooth = ls.Smooth()
	# format marker style
	# chart cannot be a 3D Chart (-4101) 
	if graphStyle.markerStyle != None and chartType != -4101:
		ms = graphStyle.markerStyle
		if ms.markerType != None:
			series.MarkerStyle = ms.MarkerType()
		if ms.markerSize != None:
			series.MarkerSize = ms.MarkerSize()
		if ms.markerColor != None:
			series.Format.Fill.Solid
			series.MarkerBackgroundColor = ms.MarkerColor()
		if ms.markerBorderColor != None:
			series.Format.Fill.Solid
			series.MarkerForegroundColor = ms.MarkerBorderColor()
		else:
			# -2 is used to set foreground to none/ no border
			series.MarkerForegroundColor = -2
	else:
		# assign xlMarkerStyleNone
		series.MarkerStyle = -4142
	return True

if runMe:
	message = None
	try:
		xlApp = LiveStream()
		errorReport = None
		wb = xlApp.ActiveWorkbook
		if sheetName == None:
			ws = xlApp.ActiveSheet
		else:
			ws = xlApp.Sheets(sheetName)
		# get chart size and location from range
		origin = ws.Cells(bb.xlRange(size)[1], bb.xlRange(size)[0])
		extent = ws.Cells(bb.xlRange(size)[3], bb.xlRange(size)[2])
		left = GetWidthHeight(origin, extent, ws)[0]
		top = GetWidthHeight(origin, extent, ws)[1]
		width = GetWidthHeight(origin, extent, ws)[2]
		height = GetWidthHeight(origin, extent, ws)[3]
		# get existing chart with same name or create new
		if ws.ChartObjects().Count > 0:
			for i in range(1, ws.ChartObjects().Count + 1, 1):
				if ws.ChartObjects().Item(i).Name == title:
					chartObject = ws.ChartObjects().Item(i)
		else:
			chartObjects = ws.ChartObjects()
			chartObject = chartObjects.Add(left, top, width, height)
			if title != None:
				chartObject.Name = title
			else:
				chartObject.Name = "Untitled"
		# update chart size
		if chartObject.Left != left:
			chartObject.Left = left
		if chartObject.Top != top:
			chartObject.Top = top
		if chartObject.Width != width:
			chartObject.Width = width
		if chartObject.Height != height:
			chartObject.Height = height
		# get chart object
		xlChart = chartObject.Chart
		# set chart type
		xlChart.ChartType = chartType
		# set chart data source range
		dataOrigin = ws.Cells(bb.xlRange(dataRange)[1], bb.xlRange(dataRange)[0])
		dataExtent = ws.Cells(bb.xlRange(dataRange)[3], bb.xlRange(dataRange)[2])
		xlChart.SetSourceData(ws.Range[dataOrigin, dataExtent])
		# set chart title
		if title != None:
			xlChart.HasTitle = True
			xlChart.ChartTitle.Text = title
		else:
			xlChart.HasTitle = False

		#########################
		### Legend Formatting ###
		#########################

		if xlChart.HasLegend:
			xlChart.Legend.Clear()
		# set text style for legend
		if legendStyle != None:
			xlChart.HasLegend = True
			# set legend box position
			if legendStyle.position != None:
				xlChart.Legend.Position = legendStyle.Position()
			# set legend text style
			if legendStyle.textStyle != None:
				FormatText(xlChart.Legend.Font, legendStyle.textStyle)
			# set border style for legend
			if legendStyle.borderStyle != None:
				FormatBorder(xlChart.Legend.Border, legendStyle.borderStyle)
			#set fill style for legend
			if legendStyle.fillStyle != None:
				FormatFill(xlChart.Legend.Interior, legendStyle.fillStyle)
			# change default Legend labels to range
			if legendStyle.labels != None:
				labels = legendStyle.Labels()
				chartSeries = xlChart.Seriescollection(1)
				catOrigin = ws.Cells(labels[1], labels[0])
				catExtent = ws.Cells(labels[3], labels[2])
				chartSeries.XValues = ws.Range[catOrigin, catExtent]

		########################
		### Chart Formatting ###
		########################

		if chartStyle != None:
			if chartStyle.borderStyle != None:
				xlChart.ChartArea.Format.Line.Visible = True
				FormatBorder(xlChart.ChartArea.Border, chartStyle.borderStyle)
			else:
				xlChart.ChartArea.Format.Line.Visible = False
			if chartStyle.fillStyle != None:
				xlChart.ChartArea.Fill.Visible = True
				xlChart.PlotArea.Fill.Visible = True
				FormatFill(xlChart.ChartArea.Interior, chartStyle.fillStyle)
			else:
				xlChart.ChartArea.Fill.Visible = False
				xlChart.PlotArea.Fill.Visible = False
			if chartStyle.roundCorners != None:
				xlChart.ChartArea.RoundedCorners = chartStyle.RoundCorners()

		########################
		### Graph Formatting ###
		########################
		
		if graphStyle != None:
			if isinstance(graphStyle, list):
				for i in range(1, len(graphStyle) + 1, 1):
					gs = graphStyle[i-1]
					xlChart.SeriesCollection(i).ClearFormats()
					s = xlChart.SeriesCollection(i)
					ApplyGraphStyle(chartType, s, gs)
			else:
				ApplyGraphStyle(chartType, xlChart.SeriesCollection(1), graphStyle)

		xlApp.ScreenUpdating = True
	except:
		xlApp.Quit()
		Marshal.ReleaseComObject(xlApp)
		# if error accurs anywhere in the process catch it
		import traceback
		errorReport = traceback.format_exc()
else:
	errorReport = "RunMe is set to False. Please set RunMe to True to create/update chart."

if errorReport == None:
	OUT = "Success!"
else:
	OUT = errorReport
