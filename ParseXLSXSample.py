import json
import os
import yaml
import math
import numpy as np
from math import cos,sin,pi,sqrt
from numpy import deg2rad, rad2deg
import openpyxl
from chemformula import ChemFormula
import inspect

def UpdateLocalSampleTable(Username,IP,Port):
	LocalPath=os.path.dirname(inspect.getfile(inspect.currentframe()))
	os.system("scp -P %d %s@%s:/RAID1/SAMPLE_information/samples_standard_boxes_info.xlsx %s" % (Port,Username,IP,LocalPath))

def TestNumbersFromTable(value):
	result=[]
	if value.find('-') > 0:
		splitted=value.split('-')
		for i in range(int(splitted[0]),int(splitted[1])+1):
			result.append(int(i))
		return result
	else:
		splitted=value.split(',')
		for i in splitted:
			result.append(int(i))
		return result

def ReadSampleFile(filename=""):
	if len(filename)==0:
		LocalPath=os.path.dirname(inspect.getfile(inspect.currentframe()))+"/samples_standard_boxes_info.xlsx"
		print(LocalPath)
		if os.path.isfile(LocalPath):
			filename=LocalPath
	workbook = openpyxl.load_workbook(filename,data_only=True)
	worksheet = workbook['SampleBox']
	dict_res={}
	for row in range(2, worksheet.max_row + 1):
		dict_row={}
		key = str(worksheet.cell(row, 1).value)
		if key:
			if key.isnumeric():
				#print(key)
				dict_row['Sample']=worksheet.cell(row, 2).value
				dict_row['Thickness']=worksheet.cell(row, 3).value
				dict_row['Mass']=float(worksheet.cell(row, 4).value)
				dict_row['Density']=worksheet.cell(row, 5).value
				dict_row['MolarMass']=0
				dict_row['SourceCoordinates']=[0,0,0] # координаты источника относительно центра системы
				dict_row['SampleCoordinates']=[0,0,-7.5]
				PosZ_str=str(worksheet.cell(row, 7).value)
				PosY_str=str(worksheet.cell(row, 8).value)
				if PosZ_str.find('34')>0:#образец стоял далеко от генератора
					dict_row['SourceCoordinates'][0]=55.5
					dict_row['SampleCoordinates'][0]=386-dict_row['Thickness']/2
				else:
					dict_row['SourceCoordinates'][0]=6+dict_row['Thickness']+1.5
					dict_row['SampleCoordinates'][0]=338+dict_row['Thickness']/2
				PosY_str=PosY_str.replace(' move box','')
				dict_row['SourceCoordinates'][2]=float(PosY_str)
				#dict_row['PositionZ']=worksheet.cell(row, 3).value
				runs=TestNumbersFromTable(str(worksheet.cell(row, 6).value))
				#print(dict_row['Sample'])
				if not (dict_row['Sample']=='Bskgr'):
					formula=ChemFormula(dict_row['Sample'])
					dict_row['MolarMass']=formula.formula_weight
					dict_row['Elements']=[]
					dict_row['NAtoms']=[]
					for i in formula.element.keys():
						dict_row['Elements'].append(i)
						dict_row['NAtoms'].append(formula.element[i])
					#print(formula.element,formula.formula_weight)
				for i in runs:
					dict_res[i]=dict_row
	return dict_res
