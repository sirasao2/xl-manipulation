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

	pt = xw.Book(preload_path)
	vm_type = pt.sheets[4].range('B7').value

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
		#print("TYPE:", vm_type)
		if k == vm_type:
			vm_count = int(v)

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
	pt = xw.Book(preload_path)
	vm_type = pt.sheets[4].range('B7').value
	#print("VM TYPE: ", vm_type)

	# find ENV type
	pt = xw.Book(build_plan_path)
	env_type = pt.sheets[5].range('C15').value

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
		if col_names[col_num] == "vf-module-model-name-base":
			col_ref_vfmmnb = col_num

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

	# creates dict of vm-types and vf-module-model-name-base
	vf_module_model_name_base_dict = {}
	for i in range(5, bp.nrows):
		vm = bp.cell_value(i, col_ref_vmt)
		modules_model_base = bp.cell_value(i, col_ref_vfmmnb)
		vf_module_model_name_base_dict[vm] = modules_model_base

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
			wb.sheets[1].range('C6').value = v + count # proper name # SUFFIX

	wb = xw.Book(preload_path)
	for k, v in vf_module_model_name_dict.items():
		if k == vm_type:
			wb.sheets[1].range('C8').value = v # proper name

	# update vnf-name
	wb = xw.Book(preload_path)
	for k, v in vnf_name_dict.items():
		if k == vm_type:
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

	# create dict for network role and subnet_name
	subnet_dict = {}
	for i in range(5, bp.nrows):
		network_role = bp.cell_value(i, 1)
		subnet_name = bp.cell_value(i, 6)
		subnet_dict[network_role] = subnet_name

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

	# implement changes to template
	for k, v in subnet_dict.items():
		for i in range(networks_sheet.nrows):
			if(networks_sheet.cell_value(i, 1) == k and k != ''):
				wb = xw.Book(preload_path)
				wb.sheets[3].range('F' + str(i+1)).value = v

def change_tag(preload_path, build_plan_path):
	"""
	This function:
		- initiates changes for all tag values
	"""
	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'Common Parameters')

	params = []
	for i in range(0, bp.nrows):
		if bp.cell_value(i, 0) == "Parameter Name":
			params.append(i)
	start = int(params[-1]) + 1

	# create dict of parameter name and associated value
	tag_dict = {}
	for i in range(start, bp.nrows):
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
	# check vm-type
	pt = xw.Book(preload_path)
	vm_type = pt.sheets[4].range('B7').value

	# grab values for vm-name and calculate appropriate suffix
	vnf_name_general = pt.sheets[1].range('C12').value

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
		if col_names[col_num] == "vm-type":
			col_ref_vmt = col_num
		if col_names[col_num] == "VFC ID (ppp)":
			col_ref_ppp = col_num

	# creates dict of vm_names and ppp
	ppp_dict = {}
	for i in range(5, bp.nrows):
		vm_types = bp.cell_value(i, col_ref_vmt)
		ppp = bp.cell_value(i, col_ref_ppp)
		ppp_dict[vm_types] = ppp

	# instantiate replacements
	for k, v in ppp_dict.items():
		if vm_type == k:
			if int(count) < 10:
				vm_name_val = vnf_name_general + v + "00" + count
			else:
				vm_name_val = vnf_name_general + v + "0" + count

	wb = xw.Book(preload_path)
	wb.sheets[4].range('C7').value = vm_name_val

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
		if k == vm_name_value:
			wb = xw.Book(preload_path)
			wb.sheets[2].range('B6').value = v

def change_vm_network_ips(preload_path, build_plan_path):
	"""
	This function:
			- fills proper information into VM-network-IPs
	"""
	# extract vm-name
	wb = xw.Book(preload_path)
	vm_name = wb.sheets[4].range('C7').value

	# open workbook and specify which sheet you would like to access
	wb = xlrd.open_workbook(build_plan_path)
	sheet_names = wb.sheet_names()
	bp = wb.sheet_by_name(u'VM-Layout')

	# string search VM-Layout column headers and assign each columns reference position (int) to a variable
	# this avoids hard coding the position of certain columns
	# order does not matter
	col_names = []
	for cols in bp.row(4): # 0 indexed, gives all column names from row 5
		col_names.append(cols.value)
	for col_num in range(0, len(col_names)):
		if col_names[col_num] == "vm-name":
			col_ref_vmn = col_num
		if col_names[col_num] == "oam_protected":
			col_ref_oam = col_num

	# create dictionary 
	oam_dict = {}
	for i in range(5, bp.nrows):
		az = bp.cell_value(i, col_ref_vmn)
		oam = bp.cell_value(i, col_ref_oam)
		oam_dict[az] = oam

	# replace values
	for k, v in oam_dict.items():
		if v == None:
			pass
		else:
			if k == vm_name:
				wb = xw.Book(preload_path)
				wb.sheets[6].range('D7').value = v

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
		if col_names[col_num] == "VFC ID (ppp)":
			col_ref_ppp = col_num
		if col_names[col_num] == "vm-type":
			col_ref_vmt = col_num

	
	# creates lists of all the names
	prb_list = []
	qrt_list = []
	lba_list = []
	vlbeph_list = []
	vlb_list = []
	ana_list = []
	msr_list = []
	cdp_list = []
	aku_list = []
	mbm_list = []
	ttn_list = []
	mgu_list = []
	con_list = []
	qtp_list = []
	ccm_list = []
	qlb_list = []
	gdn_list = []
	dbm_list = []
	akf_list = []
	dtl_list = []
	ssr_list = []
	vdb_list = []
	log_list = []
	imm_list = []
	srp_list = []
	crp_list = []
	shd_list = []
	ldr_list = []
	cgw_list = []
	dmn_list = []
	agw_list = []


	for i in range(5, bp.nrows):
		if "prb" in bp.cell_value(i, col_ref_vmn):
			prb_list.append(bp.cell_value(i, col_ref_vmn))
		if "qrt" in bp.cell_value(i, col_ref_vmn):
			qrt_list.append(bp.cell_value(i, col_ref_vmn))
		if "lba" in bp.cell_value(i, col_ref_vmn):
			lba_list.append(bp.cell_value(i, col_ref_vmn))
		if "vlb" in bp.cell_value(i, col_ref_vmn):
			vlb_list.append(bp.cell_value(i, col_ref_vmn))
		if "ana" in bp.cell_value(i, col_ref_vmn):
			ana_list.append(bp.cell_value(i, col_ref_vmn))
		if "msr" in bp.cell_value(i, col_ref_vmn):
			msr_list.append(bp.cell_value(i, col_ref_vmn))
		if "cdp" in bp.cell_value(i, col_ref_vmn):
			cdp_list.append(bp.cell_value(i, col_ref_vmn))
		if "aku" in bp.cell_value(i, col_ref_vmn):
			aku_list.append(bp.cell_value(i, col_ref_vmn))
		if "mbm" in bp.cell_value(i, col_ref_vmn):
			mbm_list.append(bp.cell_value(i, col_ref_vmn))
		if "ttn" in bp.cell_value(i, col_ref_vmn):
			ttn_list.append(bp.cell_value(i, col_ref_vmn))
		if "mgu" in bp.cell_value(i, col_ref_vmn):
			mgu_list.append(bp.cell_value(i, col_ref_vmn))
		if "con" in bp.cell_value(i, col_ref_vmn):
			con_list.append(bp.cell_value(i, col_ref_vmn))
		if "qtp" in bp.cell_value(i, col_ref_vmn):
			qtp_list.append(bp.cell_value(i, col_ref_vmn))
		if "ccm" in bp.cell_value(i, col_ref_vmn):
			ccm_list.append(bp.cell_value(i, col_ref_vmn))
		if "qlb" in bp.cell_value(i, col_ref_vmn):
			qlb_list.append(bp.cell_value(i, col_ref_vmn))
		if "gdn" in bp.cell_value(i, col_ref_vmn):
			gdn_list.append(bp.cell_value(i, col_ref_vmn))
		if "dbm" in bp.cell_value(i, col_ref_vmn):
			dbm_list.append(bp.cell_value(i, col_ref_vmn))
		if "akf" in bp.cell_value(i, col_ref_vmn):
			akf_list.append(bp.cell_value(i, col_ref_vmn))
		if "dtl" in bp.cell_value(i, col_ref_vmn):
			dtl_list.append(bp.cell_value(i, col_ref_vmn))
		if "ssr" in bp.cell_value(i, col_ref_vmn):
			ssr_list.append(bp.cell_value(i, col_ref_vmn))
		if "vdb" in bp.cell_value(i, col_ref_vmn):
			vdb_list.append(bp.cell_value(i, col_ref_vmn))
		if "log" in bp.cell_value(i, col_ref_vmn):
			log_list.append(bp.cell_value(i, col_ref_vmn))
		if "imm" in bp.cell_value(i, col_ref_vmn):
			imm_list.append(bp.cell_value(i, col_ref_vmn))
		if "srp" in bp.cell_value(i, col_ref_vmn):
			srp_list.append(bp.cell_value(i, col_ref_vmn))
		if "crp" in bp.cell_value(i, col_ref_vmn):
			crp_list.append(bp.cell_value(i, col_ref_vmn))
		if "shd" in bp.cell_value(i, col_ref_vmn):
			shd_list.append(bp.cell_value(i, col_ref_vmn))
		if "ldr" in bp.cell_value(i, col_ref_vmn):
			ldr_list.append(bp.cell_value(i, col_ref_vmn))
		if "cgw" in bp.cell_value(i, col_ref_vmn):
			cgw_list.append(bp.cell_value(i, col_ref_vmn))
		if "dmn" in bp.cell_value(i, col_ref_vmn):
			dmn_list.append(bp.cell_value(i, col_ref_vmn))
		if "agw" in bp.cell_value(i, col_ref_vmn):
			agw_list.append(bp.cell_value(i, col_ref_vmn))

	# removes brackets and white spaces
	prb_list = ('[%s]' % ','.join(map(str, prb_list)))[1:-1]
	qrt_list = ('[%s]' % ','.join(map(str, qrt_list)))[1:-1]
	lba_list = ('[%s]' % ','.join(map(str, lba_list)))[1:-1]
	vlb_list = ('[%s]' % ','.join(map(str, vlb_list)))[1:-1]
	ana_list = ('[%s]' % ','.join(map(str, ana_list)))[1:-1]
	msr_list = ('[%s]' % ','.join(map(str, msr_list)))[1:-1]
	cdp_list = ('[%s]' % ','.join(map(str, cdp_list)))[1:-1]
	aku_list = ('[%s]' % ','.join(map(str, aku_list)))[1:-1]
	mbm_list = ('[%s]' % ','.join(map(str, mbm_list)))[1:-1]
	ttn_list = ('[%s]' % ','.join(map(str, ttn_list)))[1:-1]
	mgu_list = ('[%s]' % ','.join(map(str, mgu_list)))[1:-1]
	con_list = ('[%s]' % ','.join(map(str, con_list)))[1:-1]
	qtp_list = ('[%s]' % ','.join(map(str, qtp_list)))[1:-1]
	ccm_list = ('[%s]' % ','.join(map(str, ccm_list)))[1:-1]
	qlb_list = ('[%s]' % ','.join(map(str, qlb_list)))[1:-1]
	gdn_list = ('[%s]' % ','.join(map(str, gdn_list)))[1:-1]
	dbm_list = ('[%s]' % ','.join(map(str, dbm_list)))[1:-1]
	akf_list = ('[%s]' % ','.join(map(str, akf_list)))[1:-1]
	dtl_list = ('[%s]' % ','.join(map(str, dtl_list)))[1:-1]
	ssr_list = ('[%s]' % ','.join(map(str, ssr_list)))[1:-1]
	vdb_list = ('[%s]' % ','.join(map(str, vdb_list)))[1:-1]
	log_list = ('[%s]' % ','.join(map(str, log_list)))[1:-1]
	imm_list = ('[%s]' % ','.join(map(str, imm_list)))[1:-1]
	srp_list = ('[%s]' % ','.join(map(str, srp_list)))[1:-1]
	crp_list = ('[%s]' % ','.join(map(str, crp_list)))[1:-1]
	shd_list = ('[%s]' % ','.join(map(str, shd_list)))[1:-1]
	ldr_list = ('[%s]' % ','.join(map(str, ldr_list)))[1:-1]
	cgw_list = ('[%s]' % ','.join(map(str, cgw_list)))[1:-1]
	dmn_list = ('[%s]' % ','.join(map(str, dmn_list)))[1:-1]
	agw_list = ('[%s]' % ','.join(map(str, agw_list)))[1:-1]



	# creates a dict of the vm-type values and the above lists
	names_dict = {
					"vlbagent_eph" : lba_list, 
					"vprb" : prb_list, 
					"qrouter" : qrt_list, 
					"vlb" : vlb_list,
					"analyst" : ana_list ,
					"microsevices" : msr_list ,
					"cognoscdp" : cdp_list ,
					"kuberik_aff" : aku_list ,
					"processing_eph_aff" : mbm_list ,
					"timesten" : ttn_list ,
					"managementui" : mgu_list ,
					"conductor" : con_list ,
					"qtraceprocessing_eph_aff" : qtp_list ,
					"cognosccm" : ccm_list ,
					"qtracelb" : qlb_list ,
					"gaurdian" : gdn_list ,
					"schemamanager" : dbm_list ,
					"kafka" : akf_list ,
					"distributedlock" : dtl_list ,
					"scheduledservices" : ssr_list ,
					"vertica_multi_aff" : vdb_list ,
					"logfilter_eph" : log_list ,
					"rpmrepository" : srp_list ,
					"configurationrepository" : crp_list ,
					"settingsandhealthdb_eph" : shd_list ,
					"qloader" : ldr_list ,
					"cognoscgw" : cgw_list ,
					"daemon" : dmn_list ,
					"apigw" : agw_list 
				}

	# take vm type
	pt = xw.Book(preload_path)
	vm_type = pt.sheets[4].range('B7').value

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
		if col_names[col_num] == "vm-name" and not None:
			col_ref_vmn = col_num
		if col_names[col_num] == "ext_pktinternal_ip" and not None:
			col_ref_pkip_zero = col_num
		if col_names[col_num] == "pktinternal_0_ip" and not None:
			col_ref_pk0_ip = col_num
		if col_names[col_num] == "pktinternal_1_ip" and not None:
			col_ref_pk1_ip = col_num
		if col_names[col_num] == "cdr_direct_bond_ip" and not None:
			col_ref_cdrdb_ip = col_num
		if col_names[col_num] == "vfl_pktinternal_0_ip" and not None:
			col_ref_vflpkt_ip = col_num
		if col_names[col_num] == "oam_protected" and not None:
			col_ref_oam_ip = col_num
		

	# create dictionaries of vm-names and ip's
	pkt_zero_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, col_ref_vmn)
		pkt_zero_ip = bp.cell_value(i, col_ref_pkip_zero)
		pkt_zero_dict[vm_name] = pkt_zero_ip

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

	oam_dict = {}
	for i in range(5, bp.nrows):
		vm_name = bp.cell_value(i, col_ref_vmn)
		oam_ips = bp.cell_value(i, col_ref_oam_ip)
		oam_dict[vm_name] = oam_ips

	# create dictionary of dictionaries
	ip_dict = {"ext_pktinternal_ip_0" : pkt_zero_dict,"pktinternal_0_ip" : pkt0_dict , "pktinternal_1_ip" : pkt1_dict, "cdr_direct_bond_ip" : cdr_direct_dict, "vfl_pktinternal_0_ip" :  vfl_dict, "oam_protected_ips" : oam_dict}

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

preload_list = []

paths = input("Please input path to folder containing the preload templates:\n")
while(paths == ""):
		paths = input("Please input path to folder containing the preload templates:\n")
while(paths[-1:] != "\\"):
	paths = input("ERROR: Please close folder path with a slash!\n")
while(paths == ""):
		paths = input("Please input path to folder containing the preload templates:\n")

for idx, item in enumerate(os.listdir(paths)):
	preload_list.append(paths+item)

dest_folder = input("Please enter path to the destination folder for output:\n")
while(dest_folder == ""):
		dest_folder = input("Please enter path to the destination folder for output:\n")
while(dest_folder[-1:] != "\\"):
	dest_folder = input("ERROR: Please close folder path with a slash!\n")
while(dest_folder == ""):
		dest_folder = input("Please enter path to the destination folder for output:\n")
#dest_folder = r"C:\\Users\\rs623u\\Trials\\changedTrials\\"

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
			change_vm_network_ips(preload_path, build_plan_path)
			names_tag_sheet(preload_path, build_plan_path)
			tag_sheet_indexes(preload_path, build_plan_path, str(titles))
			change_ips(preload_path, build_plan_path)
			# if titles < 9:
			# 	wb.save(dest_folder + title_list[titles] + "0" + str(titles+1) + ".xlsm")
			# else:
			wb.save(dest_folder + title_list[titles] + str(titles+1) + ".xlsm")
			wb.close()


