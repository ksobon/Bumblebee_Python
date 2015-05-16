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

graphicStyle = bb.BBGraphicStyle()

if fillStyle != None:
	graphicStyle.fillStyle = fillStyle
if textStyle != None:
	graphicStyle.textStyle = textStyle
if borderStyle != None:
	graphicStyle.borderStyle = borderStyle
	
OUT = graphicStyle
