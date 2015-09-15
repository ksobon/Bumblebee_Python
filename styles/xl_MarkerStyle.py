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

markerType = IN[0]
markerSize = IN[1]
markerColor = IN[2]
markerBorderColor = IN[3]

markerStyle = bb.BBMarkerStyle()

if markerType != None:
	markerStyle.markerType = markerType
if markerSize != None:
	markerStyle.markerSize = markerSize
if markerColor != None:
	markerStyle.markerColor = markerColor
if markerBorderColor != None:
	markerStyle.markerBorderColor = markerBorderColor

#Assign your output to the OUT variable
OUT = markerStyle
