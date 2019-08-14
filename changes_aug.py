from __future__ import unicode_literals
import io
import pandas as pd
import os
import csv
import xlrd
from xlwt import Workbook
import xlsxwriter
from openpyxl import load_workbook
import numpy as np
from pandas import ExcelWriter
import openpyxl
import xlwings as xw
from openpyxl.utils.cell import get_column_letter
import re
from collections import defaultdict

def change_general(preload_path, build_plan_path):
	global final_vf_module_name
	# extract vm
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)
	# extract file number
	file_name = preload_path[30:]
	num_append = list(re.findall(r'\d+', file_name))
	num_append = num_append[-1]
	#print(num_append)

	# create dict for vm-type : vf-module-name #
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	vf_module_name_dict = {}
	for i in range(5, bp.nrows):
		vm_type = bp.cell_value(i, 1)
		modules = bp.cell_value(i, 6)
		vf_module_name_dict[vm_type] = modules
	#print(vf_module_name_dict)

	# create dict for vm-type : vf-module-model-name #

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	vf_module_model_name_dict = {}
	for i in range(5, bp.nrows):
		vm_type = bp.cell_value(i, 1)
		modules_model = bp.cell_value(i, 5)
		vf_module_model_name_dict[vm_type] = modules_model
	#print(vf_module_model_name_dict)

	# create dict for vm-type : vnf_name #

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	vnf_name_dict = {}
	for i in range(5, bp.nrows):
		vm_type = bp.cell_value(i, 1)
		vnf_name = bp.cell_value(i, 7)
		vnf_name_dict[vm_type] = vnf_name
	#print(vnf_name_dict)

	# create dict for vm-type : vnf_type #

	vnf_type_dict = {}
	for i in range(5, bp.nrows):
		vm_type = bp.cell_value(i, 1)
		vnf_type = bp.cell_value(i, 4)
		vnf_type_dict[vm_type] = vnf_type
	#print(vnf_type_dict)

	# update vf-module-name

	wb = xw.Book(preload_path)
	for k, v in vf_module_name_dict.items():
		if k == vm_type:
			#print(v+num_append)
			wb.sheets[1].range('C6').value = v+num_append # proper name

	# update vf-module-model

	wb = xw.Book(preload_path)
	for k, v in vf_module_model_name_dict.items():
		if k == vm_type:
			#print(v+num_append)
			wb.sheets[1].range('C8').value = v+num_append # proper name

	# update vnf-name

	wb = xw.Book(preload_path)
	for k, v in vnf_name_dict.items():
		if k == vm_type:
			#print(v+num_append)
			wb.sheets[1].range('C12').value = v+num_append # proper name

	# update vnf-type

	wb = xw.Book(preload_path)
	for k, v in vnf_type_dict.items():
		if k == vm_type:
			#print(v+num_append)
			wb.sheets[1].range('C13').value = v+num_append # proper name

	final_vf_module_name = wb.sheets[1].range('C6').value















# def change_az(preload_path, build_plan_path):

# def change_networks(preload_path, build_plan_path):

# def change_vms(preload_path, build_plan_path):

# def change_tag(preload_path, build_plan_path):

preload_path = r"C:\Users\rs623u\aic_changes\preloads\pltemplate_rlba01_lba_01_delete.xlsm"
build_plan_path = r"C:\Users\rs623u\aic_changes\data\RDM52e_Automation_Build_Plan-v0.6.xlsx"
change_general(preload_path, build_plan_path)
