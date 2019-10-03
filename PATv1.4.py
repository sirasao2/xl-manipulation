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

def calculate_vm_count(build_plan_path):
	"""
	This function:
		- gathers the "# of VM's" per VM type for file generation 
		- creates list of properly named titles for final output folder based off of each vm-types vf-module-name
		- used for function calls at the end of the program
	"""
	global final_vf_module_name
	global title_list
	# find vm-type of current file
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type as a variable

	# set variable to hold vm count
	global vm_count
	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	# string search VNF-Specs column headers and assign each columns reference position (int) to a variable
	# this avoids hard coding the position of certain columns
	# order does not matter
	col_names = []
	for cols in bp.row(4): # 0 indexed, gives all column names from row 5
		col_names.append(cols.value)
	for col_num in range(0, len(col_names)):
		if col_names[col_num] == "vm-type":
			col_ref_vmt = col_num
		if col_names[col_num] == "# of VM's":
			col_ref_vmc = col_num
		if col_names[col_num] == "vf-module-name":
			col_ref_vfmn = col_num

	# create dict of vm-type and each vm-types VM Count
	count_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		count = bp.cell_value(i, col_ref_vmc)
		count_dict[vm] = count

	# cast as integer as it is pulled as a string
	for k, v in count_dict.items():
		if k != '' and k == vm_type:
			vm_count = (int(v))

	# create dictionary of vm-names and vfmn for title generation  
	vfmn_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		vfmn = bp.cell_value(i, col_ref_vfmn)
		vfmn_dict[vm] = vfmn

	# give the titles based on which vm-type the current file is
	title_list = []
	for k, v in vfmn_dict.items():
		for i in range(1, vm_count+1):
			if k != '' and k == vm_type:
				title_list.append(v)

def change_general(preload_path, build_plan_path, count):
	"""
	This function:
		- initiates changes for General tab in preload template
	"""
	# find module type
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	# string search VNF-Specs column headers and assign each columns reference position (int) to a variable
	# this avoids hard coding the position of certain columns
	# order does not matter
	col_names = []
	for cols in bp.row(4): # 0 indexed, gives all column names from row 5
		col_names.append(cols.value)
	for col_num in range(0, len(col_names)):
		if col_names[col_num] == "vm-type":
			col_ref_vmt = col_num
		if col_names[col_num] == "# of VM's":
			col_ref_vmc = col_num
		if col_names[col_num] == "vf-module-name":
			col_ref_vfmn = col_num
		if col_names[col_num] == "vf-module-model-name":
			col_ref_vfmmn = col_num
		if col_names[col_num] == "vnf-name":
			col_ref_vnfn = col_num
		if col_names[col_num] == "vnf-type":
			col_ref_vnft = col_num

	# creates dict of vm-types and vf-module-names
	vf_module_name_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		modules = bp.cell_value(i, col_ref_vfmn)
		vf_module_name_dict[vm] = modules

	# creates dict of vm-types and vf-module-model-names
	vf_module_model_name_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		modules_model = bp.cell_value(i, col_ref_vfmmn)
		vf_module_model_name_dict[vm] = modules_model

	# creates dict of vm-types and vnf-names
	vnf_name_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		vnf_name = bp.cell_value(i, col_ref_vnfn)
		vnf_name_dict[vm] = vnf_name

	# creates dict of vm-type and vnf-types
	vnf_type_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		vnf_type = bp.cell_value(i, col_ref_vnft)
		vnf_type_dict[vm] = vnf_type

	# update vf-module-name
	wb = xw.Book(preload_path)
	for k, v in vf_module_name_dict.items():
		if k == vm_type:
			if int(count) < 10:
				wb.sheets[1].range('C6').value = v + "0" + count # proper name # SUFFIX
			else:
				wb.sheets[1].range('C6').value = v + count

	# update vf-module-model
	wb = xw.Book(preload_path)
	for k, v in vf_module_model_name_dict.items():
		if k == vm_type:
			wb.sheets[1].range('C8').value = v # proper name

	# update vnf-name
	wb = xw.Book(preload_path)
	for k, v in vnf_name_dict.items():
		if k == vm_type:
			#print(v+num_append)
			wb.sheets[1].range('C12').value = v # proper name

	# update vnf-type
	wb = xw.Book(preload_path)
	for k, v in vnf_type_dict.items():
		if k == vm_type:
			wb.sheets[1].range('C13').value = v # proper name

def change_networks(preload_path, build_plan_path):
	"""
	This function:
		- initiates changes for Networks information
	"""
	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'Networks')

	# creates dict of network_role and network_name
	# these column references are hard coded
	net_dict = {}
	for i in range(5, bp.nrows):
		network_role = bp.cell_value(i, 1)
		network_name = bp.cell_value(i, 5)
		net_dict[network_role] = network_name

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	networks_sheet = wb.sheet_by_name(u'Networks')

	# implement changes to template
	for k, v in net_dict.items():
		for i in range(networks_sheet.nrows):
			if(networks_sheet.cell_value(i, 1) == k and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[3].range('C' + str(i+1)).value = v

def change_tag(preload_path, build_plan_path):
	"""
	This function:
		- initiates changes for all tag values
	"""
	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'Common Parameters')

	# create dict of parameter name and associated value
	tag_dict = {}
	for i in range(13, bp.nrows):
		par = bp.cell_value(i, 0)
		val = bp.cell_value(i, 1)
		tag_dict[par] = val

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	# implement changes to template
	for k, v in tag_dict.items():
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1) == k and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = v

def change_vm(preload_path, build_plan_path, count):
	"""
	This function:
		- initates changes for VM's tab
	"""
	# open workbook and specify which sheet you would like to access
	wb = xw.Book(preload_path)
	# grab values for vm-name and calculate appropriate suffix
	vnf_name_general = wb.sheets[1].range('C12').value
	if int(count) < 10:
		vm_name_replace = vnf_name_general + "upt00" + count
	else:
		vm_name_replace = vnf_name_general + "upt0" + count

	# parse and make replacements
	wb = xw.Book(preload_path)
	wb.sheets[4].range('C7').value = vm_name_replace

def change_az(preload_path, build_plan_path):
	"""
	This function:
		- initiates changes for AZ's  
	"""
	# open workbook and specify which sheet you would like to access
	# save vm_name
	wb = xw.Book(preload_path)
	vm_name_value = wb.sheets[4].range('C7').value

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')

	# string search VM-Layout column headers and assign each columns reference position (int) to a variable
	# this avoids hard coding the position of certain columns
	# order does not matter
	col_names = []
	for cols in bp.row(4):
		col_names.append(cols.value)
	for col_num in range(0, len(col_names)):
		if col_names[col_num] == "vm-name":
			col_ref_vmn = col_num
		if col_names[col_num] == "AZ:Compute":
			col_ref_azc = col_num

	# creates dict of vm_names and az's
	az_dict = {}
	for i in range(5, bp.nrows):
		vm_names = bp.cell_value(i, col_ref_vmn)
		az = bp.cell_value(i, col_ref_azc)
		az_dict[vm_names] = az

	# instantiates changes based on key, replaces cell with value
	for k, v in az_dict.items():
		if k == vm_name_value and k != '':
			wb = xw.Book(preload_path)
			wb.sheets[2].range('B6').value = v

def names_tag_sheet(preload_path, build_plan_path):
	"""
	This function:
		- creates the list of comma seperated names for Tag-values sheet
	"""
	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')

	# string search VM-Layout column headers and assign each columns reference position (int) to a variable
	# this avoids hard coding the position of certain columns
	# order does not matter
	col_names = []
	for cols in bp.row(4):
		col_names.append(cols.value)
	for col_num in range(0, len(col_names)):
		if col_names[col_num] == "vm-name":
			col_ref_vmn = col_num
	
	# creates lists of all the names
	prb_list = []
	qrt_list = []
	lba_list = []
	for i in range(5, bp.nrows):
		if "prb" in bp.cell_value(i, col_ref_vmn):
			prb_list.append(bp.cell_value(i, col_ref_vmn))
		if "qrt" in bp.cell_value(i, col_ref_vmn):
			qrt_list.append(bp.cell_value(i, col_ref_vmn))
		if "lba" in bp.cell_value(i, col_ref_vmn):
			lba_list.append(bp.cell_value(i, col_ref_vmn))

	# removes brackets and white spaces
	prb_list = ('[%s]' % ','.join(map(str, prb_list)))[1:-1]
	qrt_list = ('[%s]' % ','.join(map(str, qrt_list)))[1:-1]
	lba_list = ('[%s]' % ','.join(map(str, lba_list)))[1:-1]

	# creates a dict of the vm-type values and the above lists
	names_dict = {"vlbagent_eph" : lba_list, "vprb" : prb_list, "qrouter" : qrt_list}

	# take vm type
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	# search for vm type + "_names" and replace with the proper list from above
	for i in range(tag_sheet.nrows):
		for k, v in names_dict.items():
			if(tag_sheet.cell_value(i, 1) == (k + "_names") and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = str(v)

def tag_sheet_indexes(preload_path, build_plan_path, count):
	"""
	This function:
		- calculates the index values for the Tag-values sheet
	"""
	# takes file number and decrements by one
	file_name = preload_path[30:]
	num_append = list(re.findall(r'\d+', file_name))
	num_append = (int(num_append[-1]) - 1)

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	# instantiate changes to template by iterating through all rows and replacing keys with corresponding values
	for i in range(tag_sheet.nrows):
		if(tag_sheet.cell_value(i, 1).endswith("index")):
			wb = xw.Book(preload_path)
			wb.sheets[8].range('C' + str(i+1)).value = str(count)

def change_ips(preload_path, build_plan_path):
	"""
	This function:
		- initiates changes for all ip related cells in Tag-Values sheet
	"""
	# open workbook and specify which sheet you would like to access
	# save vm_name
	wb = xw.Book(preload_path)
	vm_name_value = wb.sheets[4].range('C7').value

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')

	# string search VM-Layout column headers and assign each columns reference position (int) to a variable
	# this avoids hard coding the position of certain columns
	# order does not matter
	col_names = []
	for cols in bp.row(4):
		col_names.append(cols.value)
	for col_num in range(0, len(col_names)):
		if col_names[col_num] == "vm-name":
			col_ref_vmn = col_num
		if col_names[col_num] == "pktinternal_0_ip":
			col_ref_pk0_ip = col_num
		if col_names[col_num] == "pktinternal_1_ip":
			col_ref_pk1_ip = col_num
		if col_names[col_num] == "cdr_direct_bond_ip":
			col_ref_cdrdb_ip = col_num
		if col_names[col_num] == "vfl_pktinternal_0_ip":
			col_ref_vflpkt_ip = col_num

	# create dictionaries of vm-names and ip's
	pkt0_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, col_ref_vmn)
		pk0_ip = bp.cell_value(i, col_ref_pk0_ip)
		pkt0_dict[vm_name] = pk0_ip

	pkt1_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, col_ref_vmn)
		pk1_ip = bp.cell_value(i, col_ref_pk1_ip)
		pkt1_dict[vm_name] = pk1_ip
	
	cdr_direct_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, col_ref_vmn)
		cdr = bp.cell_value(i, col_ref_cdrdb_ip)
		cdr_direct_dict[vm_name] = cdr

	vfl_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, col_ref_vmn)
		vfl = bp.cell_value(i, col_ref_vflpkt_ip)
		vfl_dict[vm_name] = vfl

	# create dictionary of dictionaries
	ip_dict = {"pktinternal_0_ip" : pkt0_dict , "pktinternal_1_ip" : pkt1_dict, "cdr_direct_bond_ip" : cdr_direct_dict, "vfl_pktinternal_0_ip" :  vfl_dict}
	
	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	# instantiate changes to template by iterating through all rows and replacing keys with corresponding values
	for i in range(tag_sheet.nrows):
		for k, v in ip_dict.items():
			if k in tag_sheet.cell_value(i, 1) and k != "":
				for k1, v1 in v.items():
					if k1 == vm_name_value:
						wb = xw.Book(preload_path)
						wb.sheets[8].range('C' + str(i+1)).value = v1


print("Hello! Meet PAT. The Preload Automation Tool!")

build_plan_path = input("Please input entire path to the build plan:\n")
while(build_plan_path == ""):
	build_plan_path = input("Please input entire path to the build plan:\n")
#build_plan_path = r"C:\Users\rs623u\Downloads\RDM52e_Automation_Build_Plan-v1.0.xlsx"

preload_list = []

paths = input("Please input path to folder containing the preload templates:\n")
while(paths == ""):
		paths = input("Please input path to folder containing the preload templates:\n")
while(paths[-1:] != "\\"):
	paths = input("ERROR: Please close folder path with a slash!\n")
while(paths == ""):
		paths = input("Please input path to folder containing the preload templates:\n")
# preload_path = r"C:\Users\rs623u\automation\VPMS-PreLoads-NC-RDM52b-Yankee-11.04.02.000.03-108202-v05\zrdm52brprb01_prb_1.xlsm"

for idx, item in enumerate(os.listdir(paths)):
	preload_list.append(paths+item)

dest_folder = input("Please enter path to the destination folder for output:\n")
while(dest_folder == ""):
		dest_folder = input("Please enter path to the destination folder for output:\n")
while(dest_folder[-1:] != "\\"):
	dest_folder = input("ERROR: Please close folder path with a slash!\n")
while(dest_folder == ""):
		dest_folder = input("Please enter path to the destination folder for output:\n")


for preload_path in preload_list:
	if "base" not in preload_path:
		calculate_vm_count(build_plan_path)
		for titles in range(0, len(title_list)):
			wb = xw.Book(preload_path)
			change_general(preload_path, build_plan_path, str(titles + 1))
			change_networks(preload_path, build_plan_path)
			change_tag(preload_path, build_plan_path)
			change_vm(preload_path, build_plan_path, str(titles + 1))
			change_az(preload_path, build_plan_path)
			names_tag_sheet(preload_path, build_plan_path)
			tag_sheet_indexes(preload_path, build_plan_path, str(titles))
			change_ips(preload_path, build_plan_path)
			if titles < 9:
				wb.save(dest_folder + title_list[titles] + "0" + str(titles+1) + ".xlsm")
			else:
				wb.save(dest_folder + title_list[titles] + str(titles+1) + ".xlsm")
			wb.close()


# PASTE TESTING
# C:\Users\rs623u\automation\data\RDM52e_Automation_Build_Plan-v0.9.xlsx
# C:\Users\rs623u\automation\preloads\
# C:\Users\rs623u\automation\changed\
# #     
# Please input entire path to the build plan:
# C:\Users\rs623u\automation\data\WHP3a_Automation_Template_v1.xlsx
# Please input path to folder containing the preload templates:
# C:\Users\rs623u\automation\Tran_Input\
# Please enter path to the destination folder for output:
# C:\Users\rs623u\automation\preloads\pltemplate_rprb01_prb_1.xlsm
#wb.save(r"C:\\Users\\rs623u\\aic_changes\\changed\\" + final_vf_module_name + ".xlsm")


# C:\Users\rs623u\automation\AIC_IMS_CP_RDM6b_CP03_Automation_Build_Plan-v1.0.xlsx
# C:\Users\rs623u\automation\vProbe_FE_vSCC_preload_11_3_7_r1_updated\


# C:\Users\rs623u\Downloads\RDM52e_Automation_Build_Plan-v1.0.xlsx
# C:\Users\rs623u\automation\VPMS-PreLoads-NC-RDM52b-Yankee-11.04.02.000.03-108202-v05\
# C:\Users\rs623u\automation\changed\

# C:\Users\rs623u\Downloads\RDM52e_Automation_Build_Plan-v1.0.xlsx

# C:\Users\rs623u\automation\RDM52c_Automation_Build_Plan_v1.0.xlsx
# C:\Users\rs623u\automation\preloads52c
# C:\Users\rs623u\automation\changed\
