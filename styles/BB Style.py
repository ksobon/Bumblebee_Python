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

sheetName = IN[0]
cellRange = IN[1]
graphicStyle = IN[2]

# Make BBStyle object if list or make multiple BBStyle objects if
# list depth == 3
if isinstance(sheetName, list):
	if isinstance(cellRange, list):
		dataObjectList = []
		for i, j, k in zip(sheetName, cellRange, graphicStyle):
			dataObjectList.append(bb.MakeStyleObject(i, j, k))
	else:
		dataObjectList = []
		for i, j in zip(sheetName, graphicStyle):
			dataObjectList.append(bb.MakeStyleObject(i,None,j))
else:
	dataObjectList = bb.MakeStyleObject(sheetName, cellRange, graphicStyle)

#Assign your output to the OUT variable
OUT = dataObjectList
