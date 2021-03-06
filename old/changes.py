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

#wb.save("C:\\Users\\rs623u\\vnf_changes_june\\changed\\" + final_vf_module_name + ".xlsm")
#wb.close()

def changes_general_vf_module_name(build_plan_path, preload_path):
	global final_vf_module_name
	# find module type
	# find file number to append
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)
	file_name = preload_path[30:]
	num_append = str(re.search(r'\d+', file_name).group())
	num_append_last_digit = (num_append[-1])


	# make dictionary of {module type : vf-module_name}
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	mrbe = wb.sheet_by_name(sheet_names[1])

	d = {}
	for i in range(mrbe.nrows):
		module = mrbe.cell_value(i, 1)
		zones = mrbe.cell_value(i, 3)
		d[module] = zones

	# look up key value, pair and replace vf-module-name + append number
	wb = xw.Book(preload_path)
	for k, v in d.items():
		if v != '' and k == vm_type:
			#print(v+num_append_last_digit)
			wb.sheets[1].range('C6').value = v+num_append_last_digit # proper name

	final_vf_module_name = wb.sheets[1].range('C6').value

# Changes in the General Tab
def changes_general_common_parameters(build_plan_path, preload_path):
	build_plan_general_values = pd.read_excel(build_plan_path, sheet_name="Common Parameters", usecols = 'B') # fetch values from build plan
	build_plan_general_values_list = build_plan_general_values.iloc[:, 0].tolist() # convert to list for indexing
	build_plan_general_values_list = [x for x in build_plan_general_values_list if str(x) != 'nan'] # remove all instances of 'nan'
	
	vnf_type_updated = build_plan_general_values_list[0] # reference updated values ...
	vnf_name_updated = build_plan_general_values_list[1]
	vnf_module_model_name_updated = build_plan_general_values_list[2]
	vnf_module_name_updated = build_plan_general_values_list[3]

	# change incorrect values on preload_path
	wb = xw.Book(preload_path) # open up a xlwings book, this is what gives you read and write permissions while preserving VBA
	wb.sheets[1].range('C13').value = vnf_type_updated
	wb.sheets[1].range('C12').value = vnf_name_updated
	wb.sheets[1].range('C8').value = vnf_module_model_name_updated
	# INCORRECT wb.sheets[1].range('C6').value = vnf_module_name_updated <- base one I need to chnage


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

	build_plan_ips_values = pd.read_excel(build_plan_path, sheet_name="IP Assignments", usecols = 'C') # fetch values from build plan
	build_plan_ips_list = build_plan_ips_values.iloc[:, 0].tolist() # convert to list for indexing
	build_plan_ips_list = [x for x in build_plan_ips_list if str(x) != 'nan']
	#print(build_plan_ips_list)

	build_plan_modules_values = pd.read_excel(build_plan_path, sheet_name="IP Assignments", usecols = 'B') # fetch values from build plan
	build_plan_modules_list = build_plan_modules_values.iloc[:, 0].tolist() # convert to list for indexing
	build_plan_modules_list = [x for x in build_plan_modules_list if str(x) != 'nan']
	#print(build_plan_modules_list)

	vips = []
	for i in range(0, len(build_plan_modules_list)):
			if build_plan_modules_list[i].endswith("VIP"):
				vips.append(build_plan_ips_list[i])
	# isolate vm type
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(sheet_names[8])

	if vm_type == 'qtracelb':
		# replace with index 1
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_floating_ip")):
				print(i)
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = vips[1]

	if vm_type == 'Vertica MC':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_floating_ip")):
				print(i)
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = vips[0]

def change_tag_values_protected_ip(build_plan_path, preload_path):
	# check the module name
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)
	# make function which appends all module ips into lists
	#ips_sheet = pd.read_excel(build_plan_path, sheet_name="IP Assignments", usecols = 'C')
	#col_C = ips_sheet.iloc[:, 0].tolist()
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	ips = wb.sheet_by_name(sheet_names[2])

	processing_list = []
	for i in range(ips.nrows):
		if(ips.cell_value(i, 0) == "processing"):
			wb = xw.Book(build_plan_path)
			processing_list.append(wb.sheets[2].range('C' + str(i+1)).value)
	processing_ips = ",".join(processing_list)	

	vertica_list = []
	for i in range(ips.nrows):
		if(ips.cell_value(i, 0) == "Vertica MC"):
			wb = xw.Book(build_plan_path)
			vertica_list.append(wb.sheets[2].range('C' + str(i+1)).value)
	vertica_ips = ",".join(vertica_list)	

	q_list = []
	for i in range(ips.nrows):
		if(ips.cell_value(i, 0) == "qtracelb"):
			wb = xw.Book(build_plan_path)
			q_list.append(wb.sheets[2].range('C' + str(i+1)).value)
	q_ips = ",".join(q_list)	

	sdn_list = []
	for i in range(ips.nrows):
		if(ips.cell_value(i, 0) == "SDN-O (Kubernete)"):
			wb = xw.Book(build_plan_path)
			sdn_list.append(wb.sheets[2].range('C' + str(i+1)).value)
	sdn_ips = ",".join(sdn_list)	

	sdn1_list = []
	for i in range(ips.nrows):
		if(ips.cell_value(i, 0) == "SDN-O (Ansible)"):
			wb = xw.Book(build_plan_path)
			sdn1_list.append(wb.sheets[2].range('C' + str(i+1)).value)
	sdn1_ips = ",".join(sdn1_list)	

	js_list = []
	for i in range(ips.nrows):
		if(ips.cell_value(i, 0) == "Jumpserver Linux"):
			wb = xw.Book(build_plan_path)
			js_list.append(wb.sheets[2].range('C' + str(i+1)).value)
	js_ips = ','.join(js_list)

	# replace protected ip cell with csv of the ips

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(sheet_names[8])

	if vm_type == 'processing':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_ips")):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = processing_ips

	if vm_type == 'Vertica MC':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_ips")):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = vertica_ips

	if vm_type == 'qtracelb':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_ips")):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = q_ips

	if vm_type == 'SDN-O (Kubernete)':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_ips")):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = sdn_ips

	if vm_type == 'SDN-O (Ansible)':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_ips")):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = sdn1_ips

	if vm_type == 'Jumpserver Linux':
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1).endswith("oam_protected_ips")):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = js_ips





# main()
#print("insert build plan path")
build_plan_path = r'C:\Users\rs623u\vnf_changes_june\builds\Build_Plan_MRBE_LSA1A_updated.xlsx'
#preload_path = r'C:\Users\rs623u\vnf_changes_june\files\zcccclrsrv01_analyst_template.xlsm'
# preload_path = r'C:\Users\rs623u\vnf_changes_june\files\zcccclrsrv01_qtracelb_template.xlsm'
# wb = xw.Book(preload_path)
# changes_general_vf_module_name(build_plan_path, preload_path)
# changes_general_common_parameters(build_plan_path, preload_path)
# changes_networks_common_parameters(build_plan_path, preload_path)
# changes_availability_zones(build_plan_path, preload_path)
# change_vm_name(build_plan_path, preload_path)
# change_vm_network_ips(build_plan_path, preload_path)
# change_tag_values_common(build_plan_path, preload_path)
# #change_tag_values_floating_ip(build_plan_path, preload_path)
# change_tag_values_protected_ip(build_plan_path, preload_path)
# wb.save("C:\\Users\\rs623u\\vnf_changes_june\\changed\\" + final_vf_module_name + ".xlsm") 
# wb.close()
# select paths


preload_list = []
print("insert path for preloads")
paths = r'C:\\Users\\rs623u\\vnf_changes_june\\files\\'
for idx, item in enumerate(os.listdir(paths)):
	preload_list.append(paths+item)

for preload_path in preload_list:
	wb = xw.Book(preload_path)
	changes_general_vf_module_name(build_plan_path, preload_path)
	changes_general_common_parameters(build_plan_path, preload_path)
	changes_networks_common_parameters(build_plan_path, preload_path)
	changes_availability_zones(build_plan_path, preload_path)
	change_vm_name(build_plan_path, preload_path)
	change_vm_network_ips(build_plan_path, preload_path)
	change_tag_values_common(build_plan_path, preload_path)
	#change_tag_values_floating_ip(build_plan_path, preload_path)
	change_tag_values_protected_ip(build_plan_path, preload_path)
	wb.save("C:\\Users\\rs623u\\vnf_changes_june\\changed\\" + final_vf_module_name + ".xlsm") 
	wb.close()

