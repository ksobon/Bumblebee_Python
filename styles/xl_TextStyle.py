# Copyright(c) 2016, David Mans, Konrad Sobon
# @arch_laboratory, http://archi-lab.net, http://neoarchaic.net

import clr
import sys

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

name = IN[0]
size = IN[1]
color = IN[2]
hAlign = IN[3]
vAlign = IN[4]
bold = IN[5]
italic = IN[6]
underline = IN[7]
strikethrough = IN[8]

textStyle = bb.BBTextStyle()

if name != None:
	textStyle.name = name
if size != None:
	textStyle.size = size
if color != None:
	textStyle.color = color
if hAlign != None:
	textStyle.horizontalAlign = hAlign
if vAlign != None:
	textStyle.verticalAlign = vAlign
if bold != None:
	textStyle.bold = bold
if italic != None:
	textStyle.italic = italic
if underline != None:
	textStyle.underline = underline
if strikethrough != None:
	textStyle.strikethrough = strikethrough

#Assign your output to the OUT variable
OUT = textStyle
