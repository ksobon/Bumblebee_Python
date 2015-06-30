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

minType = IN[0]
minValue = IN[1]
minColor = IN[2]
maxType = IN[3]
maxValue = IN[4]
maxColor = IN[5]

colorScale = bb.BB2ColorScaleFormatCondition()

if minType != None:
	colorScale.minType = minType
if minValue != None:
	colorScale.minValue = minValue
if minColor != None:
	colorScale.minColor = minColor
if maxType != None:
	colorScale.maxType = maxType
if maxValue != None:
	colorScale.maxValue = maxValue
if maxColor != None:
	colorScale.maxColor = maxColor

#Assign your output to the OUT variable
OUT = colorScale
