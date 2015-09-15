# Copyright(c) 2015, David Mans, Konrad Sobon
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
bee = bb
reload(bee)

#The inputs to this node will be stored as a list in the IN variable.
dataEnteringNode = IN

color = IN[0]
weight = IN[1]
lineType = IN[2]
compoundLineType = IN[3]
smooth = IN[4]

lineStyle = bb.BBLineStyle()

if color != None:
	lineStyle.color = color
if weight != None:
	lineStyle.weight = weight
if lineType != None:
	lineStyle.lineType = lineType
if compoundLineType != None:
	lineStyle.compoundLineType = compoundLineType
if smooth != None:
	lineStyle.smooth = smooth

#Assign your output to the OUT variable
OUT = lineStyle
