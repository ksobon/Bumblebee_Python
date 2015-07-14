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

import bumblebee as bb

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

fillStyle = IN[0]
textStyle = IN[1]
borderStyle = IN[2]
seriesName = IN[3]
value = IN[4]
percentage = IN[5]
leaderLines = IN[6]
legendKey = IN[7]
separator = IN[8]
labelPosition = IN[9]

labelStyle = bb.BBLabelStyle()

if fillStyle != None:
	labelStyle.fillStyle = fillStyle
if textStyle != None:
	labelStyle.textStyle = textStyle
if borderStyle != None:
	labelStyle.borderStyle = borderStyle
if seriesName != None:
	labelStyle.seriesName = seriesName
if value != None:
	labelStyle.value = value
if percentage != None:
	labelStyle.percentage = percentage
if leaderLines != None:
	labelStyle.leaderLines = leaderLines
if legendKey != None:
	labelStyle.legendKey = legendKey
if separator != None:
	labelStyle.separator = separator
if labelPosition != None:
	labelStyle.labelPosition = labelPosition

#Assign your output to the OUT variable
OUT = labelStyle
