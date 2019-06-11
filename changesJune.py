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

# Changes in the General Tab
def changes_general_common_parameters(build_plan_path, preload_path):
	build_plan_general_values = pd.read_excel(build_plan_path, sheet_name="Common Parameters", usecols = 'B') # fetch values from build plan
	build_plan_general_values_list = build_plan_general_values.iloc[:, 0].tolist() # convert to list for indexing
	build_plan_general_values_list = [x for x in build_plan_general_values_list if str(x) != 'nan'] # remove all instances of 'nan'
	#print(build_plan_general_values_list)
	vnf_type_updated = build_plan_general_values_list[0] # reference updated values ...
	vnf_name_updated = build_plan_general_values_list[1]
	vnf_module_model_name_updated = build_plan_general_values_list[2]
	vnf_module_name_updated = build_plan_general_values_list[3]

	# change incorrect values on preload_path
	wb = xw.Book(preload_path) # open up a xlwings book, this is what gives you read and write permissions while preserving VBA
	wb.sheets[1].range('C13').value = vnf_type_updated
	wb.sheets[1].range('C12').value = vnf_name_updated
	wb.sheets[1].range('C8').value = vnf_module_model_name_updated
	wb.sheets[1].range('C6').value = vnf_module_name_updated

def changes_networks_common_parameters(build_plan_path, path):
	build_plan_network_name_values = pd.read_excel(build_plan_path, sheet_name="Common Parameters", usecols = 'B')
	build_plan_subnet_name_values = pd.read_excel(build_plan_path, sheet_name="Common Parameters", usecols = 'C')

	build_plan_network_name_values_list = build_plan_network_name_values.iloc[:, 0].tolist() # convert to list for indexing
	build_plan_network_name_values_list = [x for x in build_plan_network_name_values_list if str(x) != 'nan'] # remove all instances of 'nan'

	build_plan_subnet_name_values_list = build_plan_subnet_name_values.iloc[:, 0].tolist() # convert to list for indexing
	build_plan_subnet_name_values_list = [x for x in build_plan_subnet_name_values_list if str(x) != 'nan'] # remove all instances of 'nan'

	oam_protected_network_name_updated = build_plan_network_name_values_list[6]
	vprobes_mgmt_network_name_updated = build_plan_network_name_values_list[7]
	cdr_direct_network_name_updated = build_plan_network_name_values_list[8]
	backend_ic_network_name_updated = build_plan_network_name_values_list[9]

	oam_protected_subnet_name_updated = build_plan_subnet_name_values_list[1]
	vprobes_mgmt_subnet_name_updated = build_plan_subnet_name_values_list[2]
	cdr_direct_subnet_name_updated = build_plan_subnet_name_values_list[3]
	backend_ic_subnet_name_updated = build_plan_subnet_name_values_list[4]

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	networks_sheet = wb.sheet_by_name(sheet_names[3])

	# update networks
	for i in range(networks_sheet.nrows):
		if(networks_sheet.cell_value(i, 1) == 'backend_ic'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('C' + str(i+1)).value = backend_ic_network_name_updated
		if(networks_sheet.cell_value(i, 1) == 'vprobes_mgmt'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('C' + str(i+1)).value = vprobes_mgmt_network_name_updated
		if(networks_sheet.cell_value(i, 1) == 'oam_protected'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('C' + str(i+1)).value = oam_protected_network_name_updated
		if(networks_sheet.cell_value(i, 1) == 'cdr_direct'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('C' + str(i+1)).value = cdr_direct_network_name_updated

	# update subnets
	for i in range(networks_sheet.nrows):
		if(networks_sheet.cell_value(i, 1) == 'backend_ic'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('F' + str(i+1)).value = backend_ic_subnet_name_updated
		if(networks_sheet.cell_value(i, 1) == 'vprobes_mgmt'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('F' + str(i+1)).value = vprobes_mgmt_subnet_name_updated
		if(networks_sheet.cell_value(i, 1) == 'oam_protected'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('F' + str(i+1)).value = oam_protected_subnet_name_updated
		if(networks_sheet.cell_value(i, 1) == 'cdr_direct'):
			wb = xw.Book(preload_path)
			wb.sheets[3].range('F' + str(i+1)).value = cdr_direct_subnet_name_updated

def changes_availability_zones(build_plan_path, preload_path):
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	file_name = preload_path[30:]
	num_append = str(re.search(r'\d+', file_name).group())

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	mrbe = wb.sheet_by_name(sheet_names[1])

	# excel to dictionary
	d = {}
	for i in range(mrbe.nrows):
		module = mrbe.cell_value(i, 1)
		zones = mrbe.cell_value(i, 4)
		d[module] = zones


	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	az = wb.sheet_by_name(sheet_names[2])

	for k, v in d.items():
		if k == vm_type:
			wb = xw.Book(preload_path) # open up a xlwings book, this is what gives you read and write permissions while preserving VBA
			wb.sheets[2].range('B6').value = v.strip()+num_append.strip()

def change_vm_name(build_plan_path, preload_path):
	# extract vm_type
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)
	# extract file number
	file_name = preload_path[30:]
	num_append = str(re.search(r'\d+', file_name).group())
	#print(num_append)

	# create dict
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	mrbe = wb.sheet_by_name(sheet_names[1])

	vm_name_dict = {}
	for i in range(mrbe.nrows):
		module = mrbe.cell_value(i, 1)
		vm_names = mrbe.cell_value(i, 9)
		vm_name_dict[module] = vm_names

	# replace c7
	for k, v in vm_name_dict.items():
		if k == vm_type:
			wb = xw.Book(preload_path)
			wb.sheets[4].range('C7').value = v.strip()+num_append.strip()

def change_vm_network_ips(build_plan_path, preload_path):
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'C')
	col_C_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_name = col_C_list_vms[-1] # save VM NAME
	
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	ips = wb.sheet_by_name(sheet_names[2])

	ips_dict = {}
	for i in range(ips.nrows):
		modules = ips.cell_value(i, 1)
		ipa = ips.cell_value(i, 2)
		ips_dict[modules] = ipa

	for k, v in ips_dict.items():
		if k  == vm_name:
			wb = xw.Book(preload_path)
			wb.sheets[6].range('D7').value = v

def change_tag_values_common(build_plan_path, preload_path):
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	common = wb.sheet_by_name(sheet_names[0])

	tag_dict = {}
	for i in range(18, common.nrows):
		par_names = (common.cell_value(i, 0))
		values = (common.cell_value(i, 1))
		tag_dict[par_names] = values

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(sheet_names[8])

	for k, v in tag_dict.items():
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1) == k):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = v

	



# main()
build_plan_path = r'C:\Users\rs623u\vnf_changes_june\files\Build_Plan_MRBE_LSA1A_updated.xlsx'
#preload_path = r'C:\Users\rs623u\vnf_changes_june\files\zcccclrsrv01_analyst_template.xlsm'
preload_path = r'C:\Users\rs623u\vnf_changes_june\files\zcccclrsrv01_qtracelb_template.xlsm'
#changes_general_common_parameters(build_plan_path, preload_path)
#changes_networks_common_parameters(build_plan_path, preload_path)
#changes_availability_zones(build_plan_path, preload_path)
#change_vm_name(build_plan_path, preload_path)
#change_vm_network_ips(build_plan_path, preload_path)
change_tag_values_common(build_plan_path, preload_path)
