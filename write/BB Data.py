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
origin = IN[1]
data = IN[2]

# Make BBData object if list or make multiple BBData objects if
# list depth == 3
if isinstance(sheetName, list):
	if isinstance(origin, list):
		dataObjectList = []
		for i, j, k in zip(sheetName, origin, data):
			dataObjectList.append(bb.MakeDataObject(i, j, k))
	else:
		dataObjectList = []
		for i, j in zip(sheetName, data):
			dataObjectList.append(bb.MakeDataObject(i,None,j))
else:
	dataObjectList = bb.MakeDataObject(sheetName, origin, data)

#Assign your output to the OUT variable
OUT = dataObjectList
