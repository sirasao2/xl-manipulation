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
	# find module type
	# find file number to append
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)
	file_name = preload_path[30:]
	num_append = list(re.findall(r'\d+', file_name))
	#print(num_append)
	num_append = num_append[-1]
	#print(num_append)


	#  dict for vm-type : vf-module-name #
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	vf_module_name_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, 2)
		modules = bp.cell_value(i, 8)
		vf_module_name_dict[vm] = modules

	# create dict for vm-type : vf-module-model-name #

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	vf_module_model_name_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, 2)
		modules_model = bp.cell_value(i, 6)
		vf_module_model_name_dict[vm] = modules_model
	#print(vf_module_model_name_dict)

	# create dict for vm-type : vnf_name #

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VNF-Specs')

	vnf_name_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, 2)
		vnf_name = bp.cell_value(i, 9)
		vnf_name_dict[vm] = vnf_name
	#print(vnf_name_dict)

	# create dict for vm-type : vnf_type #

	vnf_type_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, 2)
		vnf_type = bp.cell_value(i, 5)
		vnf_type_dict[vm] = vnf_type
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
			#print(v+num_append)
			wb.sheets[1].range('C13').value = v # proper name

	final_vf_module_name = wb.sheets[1].range('C6').value

# def change_az(preload_path, build_plan_path):

def change_networks(preload_path, build_plan_path):
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	#print(sheet_names)
	bp = wb.sheet_by_name(u'Networks')

	net_dict = {}
	for i in range(5, bp.nrows):
		network_role = bp.cell_value(i, 1)
		network_name = bp.cell_value(i, 5)
		net_dict[network_role] = network_name

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	networks_sheet = wb.sheet_by_name(u'Networks')

	for k, v in net_dict.items():
		for i in range(networks_sheet.nrows):
			if(networks_sheet.cell_value(i, 1) == k and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[3].range('C' + str(i+1)).value = v

# def change_vms(preload_path, build_plan_path):

def change_tag(preload_path, build_plan_path):
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'Common Parameters')

	tag_dict = {}
	for i in range(13, bp.nrows):
		par = bp.cell_value(i, 0)
		val = bp.cell_value(i, 1)
		tag_dict[par] = val

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	for k, v in tag_dict.items():
		for i in range(tag_sheet.nrows):
			if(tag_sheet.cell_value(i, 1) == k and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = v

def change_vm(preload_path, build_plan_path):
	file_name = preload_path[30:]
	num_append = list(re.findall(r'\d+', file_name))
	#print(num_append)
	num_append = num_append[0]
	#print(num_append)
	wb = xw.Book(preload_path)
	vnf_name_general = wb.sheets[1].range('C12').value
	vm_name_replace = vnf_name_general + "upt0" + num_append
	#print(vm_name_replace)
	wb = xw.Book(preload_path)
	wb.sheets[4].range('C7').value = vm_name_replace

def change_az(preload_path, build_plan_path):
	wb = xw.Book(preload_path)
	vm_name_value = wb.sheets[4].range('C7').value
	#print(vm_name_value)

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')

	az_dict = {}
	for i in range(5, bp.nrows):
		vm_names = bp.cell_value(i, 6)
		az = bp.cell_value(i, 7)
		az_dict[vm_names] = az

	for k, v in az_dict.items():
		if k == vm_name_value and k != '':
			wb = xw.Book(preload_path)
			wb.sheets[2].range('B6').value = v

def names_tag_sheet(preload_path, build_plan_path):
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')
	
	prb_list = []
	qrt_list = []
	lba_list = []
	for i in range(5, bp.nrows):
		if "prb" in bp.cell_value(i, 6):
			prb_list.append(bp.cell_value(i, 6))
		if "qrt" in bp.cell_value(i, 6):
			qrt_list.append(bp.cell_value(i, 6))
		if "lba" in bp.cell_value(i, 6):
			lba_list.append(bp.cell_value(i, 6))

	names_dict = {"vlbagent_eph" : lba_list, "vprb" : prb_list, "qrouter" : qrt_list}
	#print(names_dict)

	# take vm type
	extract_vm = pd.read_excel(preload_path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)
	# search for vm type + _names and replace with list
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	for i in range(tag_sheet.nrows):
		for k, v in names_dict.items():
			if(tag_sheet.cell_value(i, 1) == (k + "_names") and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[8].range('C' + str(i+1)).value = str(v)

def tag_sheet_indexes(preload_path, build_plan_path):
	file_name = preload_path[30:]
	num_append = list(re.findall(r'\d+', file_name))
	num_append = (int(num_append[-1]) - 1)

	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	for i in range(tag_sheet.nrows):
		if(tag_sheet.cell_value(i, 1).endswith("index")):
			wb = xw.Book(preload_path)
			wb.sheets[8].range('C' + str(i+1)).value = str(num_append)

def change_ips(preload_path, build_plan_path):
	wb = xw.Book(preload_path)
	vm_name_value = wb.sheets[4].range('C7').value
	#print(vm_name_value)

	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')

	pkt0_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, 6)
		pk0_ip = bp.cell_value(i, 9)
		pkt0_dict[vm_name] = pk0_ip
	#print(pkt0_dict)

	pkt1_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, 6)
		pk1_ip = bp.cell_value(i, 10)
		pkt1_dict[vm_name] = pk1_ip
	#print(pkt1_dict)
	
	cdr_direct_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, 6)
		cdr = bp.cell_value(i, 11)
		cdr_direct_dict[vm_name] = cdr
	#print(cdr_direct_dict)

	vfl_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, 6)
		vfl = bp.cell_value(i, 12)
		vfl_dict[vm_name] = vfl
	#print(vfl_dict)

	ip_dict = {"pktinternal_0_ip" : pkt0_dict , "pktinternal_1_ip" : pkt1_dict, "cdr_direct_bond_ip" : cdr_direct_dict, "vfl_pktinternal_0_ip" :  vfl_dict}
	
	# print(ip_dict)
	# print("\n")

	# for k, v in ip_dict.items():
 #   		for k1, v1 in v.items():
 #   			print("k1", k1)
 #   			print("v1", v1)
	
	wb = xlrd.open_workbook(preload_path)
	sheet_names = wb.sheet_names()
	tag_sheet = wb.sheet_by_name(u'Tag-values')

	for i in range(tag_sheet.nrows):
		for k, v in ip_dict.items():
			if k in tag_sheet.cell_value(i, 1) and k != "":
				for k1, v1 in v.items():
					if k1 == vm_name_value:
						wb = xw.Book(preload_path)
						wb.sheets[8].range('C' + str(i+1)).value = v1


#preload_path = r"C:\Users\rs623u\aic_changes\preloads\pltemplate_rlba01_lba_01_delete.xlsm"
print("Hello!")

# build_plan_path = r"C:\Users\rs623u\automation\data\RDM52e_Automation_Build_Plan-v0.9.xlsx"
# preload_path = r"C:\Users\rs623u\automation\preloads\pltemplate_rprb01_prb_1.xlsm"
#preload_path = r"C:\Users\rs623u\aic_changes\preloads\pltemplate_rlba01_lba_01_delete.xlsm"

build_plan_path = input("Please input entire path to the build plan:\n")
while(build_plan_path == ""):
	build_plan_path = input("Please input entire path to the build plan:\n")

preload_list = []

paths = input("Please input path to folder containing the preload templates:\n")
while(paths == ""):
	paths = input("Please input path to folder containing the preload templates:\n")


paths = r"C:\\Users\\rs623u\\aic_changes\\preloads\\"
for idx, item in enumerate(os.listdir(paths)):
	preload_list.append(paths+item)

dest_folder = input("Please enter path to the destination folder for output:\n")
while(dest_folder == ""):
	dest_folder = input("Please enter path to the destination folder for output:\n")

for preload_path in preload_list:
	if "base" not in preload_path:
		wb = xw.Book(preload_path)
		change_general(preload_path, build_plan_path)
		change_networks(preload_path, build_plan_path)
		change_tag(preload_path, build_plan_path)
		change_vm(preload_path, build_plan_path)
		change_az(preload_path, build_plan_path)
		names_tag_sheet(preload_path, build_plan_path)
		tag_sheet_indexes(preload_path, build_plan_path)
		change_ips(preload_path, build_plan_path)
		wb.save(dest_folder + final_vf_module_name + ".xlsm")
		wb.close()




# PASTE TESTING
#     C:\Users\rs623u\automation\data\RDM52e_Automation_Build_Plan-v0.9.xlsx
#     C:\Users\rs623u\automation\preloads\
#     C:\Users\rs623u\automation\vpms_pl\
#     C:\Users\rs623u\automation\changed\

# C:\Users\rs623u\automation\preloads\pltemplate_rprb01_prb_1.xlsm
#wb.save(r"C:\\Users\\rs623u\\aic_changes\\changed\\" + final_vf_module_name + ".xlsm")
