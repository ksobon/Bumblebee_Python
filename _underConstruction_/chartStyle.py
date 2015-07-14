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
	sys.path.append(bbPath)

import bumblebee as bb

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

fillStyle = IN[0]
textStyle = IN[1]
borderStyle = IN[2]
roundCorners = IN[3]

chartStyle = bb.BBChartStyle()

if fillStyle != None:
	chartStyle.fillStyle = fillStyle
if textStyle != None:
	chartStyle.textStyle = textStyle
if borderStyle != None:
	chartStyle.borderStyle = borderStyle
if roundCorners != None:
	chartStyle.roundCorners = roundCorners
	
OUT = chartStyle
