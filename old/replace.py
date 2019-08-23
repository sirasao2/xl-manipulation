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

def extract_common_general(build_path):
	extract = pd.read_excel(build_path, sheet_name="Common Parameters", usecols = 'B') # Extract column B from Build Path
	col_B_list = extract.iloc[:, 0].tolist() # Save values of Column B to a list
	#print("Column A:", col_A_list)

	correct_vnf_type = col_B_list[0] # Correct vnf-type, reference by index in above array
	#print("Correct vnf-type:", correct_vnf_type)
	
	correct_vnf_name = col_B_list[1] # Correct vnf-name, reference by index in above array
	#print("Correct vnf-name:", correct_vnf_name)

	wb = xw.Book(path) # open up a xlwings book, this is what gives you read and write permissions while preserving VBA
	wb.sheets[1].range('C12').value = correct_vnf_name # wb.sheets[i], where i = index_of_sheet (0 indexed); range(COLUMN-ROW); .value, isolate value, assign correct value
	wb.sheets[1].range('C13').value = correct_vnf_type # wb.sheets[i], where i = index_of_sheet (0 indexed); range(COLUMN-ROW); .value, isolate value, assign correct value

def replace_module_model_general(path):
	vf_module_model_name_dict = { # create dictionary of pairs
		"rpmrepository" :	"rpmrepository",
		"conductor" :	"VpmsBeVf1137..BE_Add_On_Module_conductor..module-19",
		"settingsandhealthdb" :	"VpmsBeVf1137..BE_Add_On_Module_settingsandhealthdb..module-3",
		"apigw" :	"VpmsBeVf1137..BE_Add_On_Module_apigw..module-23",
		"analyst" :	"VpmsBeVf1137..BE_Add_On_Module_analyst..module-24",
		"configrepo" :	"VpmsBeVf1137..BE_Add_On_Module_configurationrepository..module-18",
		"distributedlock" :	"VpmsBeVf1137..BE_Add_On_Module_distributedlock..module-16",
		"vertica" :	"VpmsBeVf1137..BE_Add_On_Module_vertica_multi_aff..module-1",
		"schemamanager" :	"VpmsBeVf1137..BE_Add_On_Module_schemamanager..module-4",
		"microservices" :	"VpmsBeVf1137..BE_Add_On_Module_microservices..module-12",
		"managementui" :	"VpmsBeVf1137..BE_Add_On_Module_managementui..module-13",
		"guardian" :	"VpmsBeVf1137..BE_Add_On_Module_guardian..module-15",
		"qloader" :	"VpmsBeVf1137..BE_Add_On_Module_qloader..module-10",
		"scheduledservices" :	"VpmsBeVf1137..BE_Add_On_Module_scheduledservices..module-5",
		"daemon" :	"VpmsBeVf1137..BE_Add_On_Module_daemon..module-17",
		"timesten" :	"VpmsBeVf1137..BE_Add_On_Module_timesten..module-2",
		"processing" :	"VpmsBeVf1137..BE_Add_On_Module_processing_eph..module-11",
		"qtraceprocessing" :	"VpmsBeVf1137..BE_Add_On_Module_qtraceprocessing_eph_aff..module-7",
		"qtracelb" :	"VpmsBeVf1137..BE_Add_On_Module_qtracelb..module-8",
		"cognoscgw" :	"VpmsBeVf1137..BE_Add_On_Module_cognoscgw..module-20",
		"cognosccm" :	"VpmsBeVf1137..BE_Add_On_Module_cognosccm..module-22",
		"cognoscdp" :	"VpmsBeVf1137..BE_Add_On_Module_cognoscdp..module-21",
		"logfilter" :	"VpmsBeVf1137..BE_Add_On_Module_logfilter_eph..module-14",
		"QTrace Server" :	"VpmsBeVf1137..BE_Add_On_Module_qtrace..module-9"
	}

	# extract vm model for verification
	extract_vm = pd.read_excel(path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	for key, value in vf_module_model_name_dict.items():
		if key == vm_type:
			wb = xw.Book(path)
			wb.sheets[1].range('C8').value = value

def replace_module_name_general(path):
	vf_module_name_dict = {
		"rpmrepository" : "zrdm5anbea01_rpmrepository_1",
		"conductor" :	"zrdm5anbea01_conductor_1 to 2",
		"settingsandhealthdb" :	"zrdm5anbea01_settingsandhealthdb_1 to 3",
		"apigw" :	"zrdm5anbea01_apigw_1 to 2",
		"analyst" :	"zrdm5anbea01_analyst_1 to 2",
		"configrepo" :	"zrdm5anbea01_configrepo_1",
		"distributedlock" :	"zrdm5anbea01_distributedlock_1 to 3",
		"vertica" :	"zrdm5anbea01_vertica_1 to 4",
		"schemamanager" :	"zrdm5anbea01_schemamanager_1",
		"microservices	" : "zrdm5anbea01_microservices_1 to 2",
		"managementuiv" :	"zrdm5anbea01_managementui_1 to 2",
		"guardian" : "zrdm5anbea01_guardian_1 to 2",
		"qloader" : "zrdm5anbea01_qloader_1 to 2",
		"scheduledservices" : "zrdm5anbea01_scheduledservices_1",
		"daemon" : "zrdm5anbea01_daemon_1",
		"timesten" : "zrdm5anbea01_timesten_1",	
		"processing" : "zrdm5anbea01_Mobility_processing_1",
		"qtraceprocessing" : "zrdm5anbea01_qtraceprocessing_1",
		"qtracelb" : "zrdm5anbea01_qtracelb_1",
		"cognoscgw" : "zrdm5anbea01_cognoscgw_1 to 2",
		"cognosccm" : "zrdm5anbea01_cognosccm_1 to 2",
		"cognoscdp" : "zrdm5anbea01_cognoscdp_1 to 2v",
		"logfilter" : "zrdm5anbea01_logfilter_1"
	}
	extract_vm = pd.read_excel(path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	for key, value in vf_module_name_dict.items():
		if key == vm_type:
			wb = xw.Book(path)
			wb.sheets[1].range('C6').value = value

def replace_networks(path):
	extract_nn = pd.read_excel(build_path, sheet_name="Common Parameters", usecols = 'B')
	col_B_list_network_names = extract_nn.iloc[:, 0].tolist() 
	#print(col_B_list_network_names)
	nn_oam_protected = col_B_list_network_names[10]
	nn_vprobes_mgmt = col_B_list_network_names[11]
	nn_cdr_direction = col_B_list_network_names[12]
	nn_backend_ic = col_B_list_network_names[13]

	# print(oam_protected)
	# print(vprobes_mgmt)
	# print(cdr_direction)
	# print(backend_ic)

	extract_ipv4 = pd.read_excel(build_path, sheet_name="Common Parameters", usecols = 'C')
	col_C_list_ipv4 = extract_ipv4.iloc[:, 0].tolist()
	#print(col_C_list_ipv4)

	ip4_oam_protected = col_C_list_ipv4[10]
	ip4_vprobes_mgmt = col_C_list_ipv4[11]
	ip4_cdr_direction = col_C_list_ipv4[12]
	ip4_backend_ic = col_C_list_ipv4[13]

	# print(ip4_oam_protected)
	# print(ip4_vprobes_mgmt)
	# print(ip4_cdr_direction)
	# print(ip4_backend_ic)
	
	# Change network name
	wb = xw.Book(path)
	if wb.sheets[3].range('B7').value is not None:
		wb.sheets[3].range('C7').value = nn_backend_ic
	if wb.sheets[3].range('B8').value is not None:
		wb.sheets[3].range('C8').value = nn_vprobes_mgmt
	if wb.sheets[3].range('B9').value is not None:
		wb.sheets[3].range('C9').value = nn_oam_protected
	if wb.sheets[3].range('B10').value is not None:
		wb.sheets[3].range('C10').value = nn_cdr_direction

	# # Change ip4-subnet-name
	if wb.sheets[3].range('B7').value is not None:
		wb.sheets[3].range('F7').value = ip4_backend_ic
	if wb.sheets[3].range('B8').value is not None:
		wb.sheets[3].range('F8').value = ip4_vprobes_mgmt
	if wb.sheets[3].range('B9').value is not None:
		wb.sheets[3].range('F9').value = ip4_oam_protected
	if wb.sheets[3].range('B10').value is not None:
		wb.sheets[3].range('F10').value = ip4_cdr_direction

def change_vm_name(path):
	vm_name_dict = {
		"rpmrepository" : "zrdm5anbea01srp00" ,
		"conductor" :	"zrdm5anbea01con00" ,
		"settingsandhealthdb" :	"zrdm5anbea01shd00" ,
		"apigw" :	"zrdm5anbea01agw00" ,
		"analyst" : "z0rdm5anbea01ana00" ,
		"configrepo" :	"zrdm5anbea01crp00" ,
		"distributedlock" : "zrdm5anbea01dtl00" , 
		"vertica" :	"zrdm5anbea01vdb00" ,
		"schemamanager" :	"zrdm5anbea01dbm00" ,
		"microservices" :	"zrdm5anbea01msr00" ,
		"managementui" :	"zrdm5anbea01mgu00 " ,
		"guardian" :	"zrdm5anbea01gdn00 " ,
		"qloader" :	"zrdm5anbea01ldr00" ,
		"scheduledservices" :	"zrdm5anbea01ssr00" ,
		"daemon" :	"zrdm5anbea01dmn00" ,
		"timesten" :	 "zrdm5anbea01ttn00" ,
		"processing" :	"zrdm5anbea01mbm00" ,
		"qtraceprocessing" : "zrdm5anbea01qtp00" ,
		"qtracelb" :	"zrdm5anbea01qlb00" ,
		"cognoscgw" :	"zrdm5anbea01cgw00" ,
		"cognosccm" :	"zrdm5anbea01ccm00" ,
		"cognoscdp" :	"zrdm5anbea01cdp00" ,
		"logfilter" :	"zrdm5anbea01log00" ,
		"QTrace Server" :	"LLCCCCRAVPMW0" ,
		"Vertica Aux DB" :	"zrdm5anbea01adb00" ,
		"Kafka" :	"zrdm5anbea01kfk00" ,
		"Vertica MC" :	"zrdm5anbea01vmc00" ,
		"SDN-O (Ansible)" :	"zrdm5anbea01sda00" ,
		"SDN-O (Kubernete)" :	"zrdm5anbea01sdk00" ,
		"Data Collection & Analysis" :	"zrdm5anbea01dca00" ,
		"Jumpserver Windows" :	"LLCCCCCRAVPMA0" ,
		"Jumpserver Linux" :	"zrdm5anbea01ljs00" 
	}

	extract_vm = pd.read_excel(path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	vm_num_appendation = (path[-6])

	for key, value in vm_name_dict.items():
		if key == vm_type:
			wb = xw.Book(path)
			wb.sheets[4].range('C7').value = value + vm_num_appendation

def change_az(path):

	dict_az = {
		"rpmrepository" :	"rdm5a-kvm-az0" ,
		"conductor" :	"rdm5a-kvm-az0",
		"settingsandhealthdb" :	"rdm5a-kvm-az0",
		"apigw" :	"rdm5a-kvm-az0",
		"analyst" :	"rdm5a-kvm-az0",
		"configrepo" : "rdm5a-kvm-az0",
		"distributedlock" :	"rdm5a-kvm-az0",
		"vertica" :	"rdm5a-kvm-az0",
		"schemamanager" :	"rdm5a-kvm-az0",
		"microservices" :	"rdm5a-kvm-az0",
		"managementui" : "rdm5a-kvm-az0",
		"guardian" :	"rdm5a-kvm-az0",
		"qloader" :	"rdm5a-kvm-az0",
		"scheduledservices" :	"rdm5a-kvm-az0",
		"daemon" :	"rdm5a-kvm-az0",
		"timesten" :	"rdm5a-kvm-az0",
		"processing" :	"rdm5a-kvm-az0",
		"qtraceprocessing" :	"rdm5a-kvm-az0",
		"qtracelb" :	"rdm5a-kvm-az0",
		"cognoscgw" :	"rdm5a-kvm-az0",
		"cognosccm" :	"rdm5a-kvm-az0",
		"cognoscdp" :	"rdm5a-kvm-az0",
		"logfilter" :	"rdm5a-kvm-az0",
		"QTrace Server" :	"rdm5a-kvm-az0",
		"Vertica Aux DB" :	"rdm5a-kvm-az0",
		"Kafka" :	"rdm5a-kvm-az0",
		"Vertica MC" :	"rdm5a-kvm-az0",
		"SDN-O (Ansible)" :	"rdm5a-kvm-az0",
		"SDN-O (Kubernete)" :  "rdm5a-kvm-az0",
		"Data Collection & Analysis" :	"rdm5a-kvm-az0",
		"Jumpserver Windows" :	"rdm5a-kvm-az0",
		"Jumpserver Linux" :	"rdm5a-kvm-az0"
	}
	vm_num = int(path[-6])
	extract_vm = pd.read_excel(path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	for key, value in dict_az.items():
		if key == vm_type:
			if vm_num % 2 == 1:
				wb = xw.Book(path)
				wb.sheets[2].range('B6').value = value + str(1)
			elif vm_num % 2 == 0:
				wb = xw.Book(path)
				wb.sheets[2].range('B6').value = value + str(2)

def change_ips(path):
	ips_dict = {
		"rpmrepository" :	"107.112.129.75",
		"conductor" :	"107.112.129.",
		"apigw" :		"107.112.129.",
		"analyst" :		"107.112.129.",
		"managementui" : "107.112.129.",
		"guardian" :	"107.112.129.",
		"processing" :	"107.112.129."	,
		"qtracelb" :	"107.112.129.",
		"cognoscgw" : "107.112.129.",
		"logfilter" :	"107.112.129.82"
	}

	vm_num = int(path[-6])
	extract_vm = pd.read_excel(path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type

	for key, value in ips_dict.items():
		if key == vm_type and vm_type == "rpmrepository":
			# replace
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value 
		if key == vm_type and vm_type == "logfilter":
			# replace
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value
		if key == vm_type and vm_type == "conductor" and vm_num == 1:
			# replace and append 90
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(90)
		elif vm_num == 2:
			# replace and append 93
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(93)
		if key == vm_type and vm_type == "apigw" and vm_num == 1:
			# replace and append 85
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(85)
		elif vm_num == 2:
			# replace and append 86
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(86)
		if key == vm_type and vm_type == "analyst" and vm_num == 1:
			# replace and append 96
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(96)
		elif vm_num == 2:
			# replace and append 97
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(97)
		if key == vm_type and vm_type == "managementui" and vm_num == 1:
			# replace and append 99
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(99)
		elif vm_num == 2:
			# replace and append 104
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(104)
		if key == vm_type and vm_type == "guardian" and vm_num == 1:
			# ip incomplete
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value 
		elif vm_num == 2:
			# ip incomplete
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value 
		if key == vm_type and vm_type == "processing" and vm_num == 1:
			# replace and append 76
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(76)
		elif vm_num == 2:
			# replace and append 77
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(77)
		if key == vm_type and vm_type == "qtracelb" and vm_num == 1:
			# replace and append 78
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(78)
		elif vm_num == 2:
			# replace and append 79
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(79)
		if key == vm_type and vm_type == "cognoscgw" and vm_num == 1:
			# replace and append 82
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(82)
		elif vm_num == 2:
			# replace and append 83
			wb = xw.Book(path)
			wb.sheets[6].range('D7').value = value + str(83)

def replace_tag_values(path):
	tags_common_dict = {
		"security_group_name" :	"VPMS-FN-26071-T-BE-01_11_3_7" ,
		"be_security_group_id" :	"f80c6c35-83ad-4bf2-a1e2-2f5f98564d5a" ,
		"domain_name" :	"novalocal" ,
		"environment_context" :	"RDM5a" ,
		"global_mtu" :	"1400" ,
		"site_name" :	"RDM5a" ,
		"tenant_name" :	"VPMS-FN-26071-T-BE-01" ,
		"workload_context" :	"RDM5a" ,
		"oam_protected_gateway" :	"107.112.128.1" ,
		"oam_protected_net_name" :	"MNS-FN-25180-T-01Shared_oam_protected_net_1" ,
		"oam_protected_route_cidrs" :	"107.112.128.0/21" ,
		"int_vprobes_mgmt_net_gateway" :	"10.25.8.1" ,
		"int_vprobes_mgmt_net_name" :	"VPMS-FN-26071-T-BE-01_com_vprobes_int_mgmt_net_1" ,
		"int_vprobes_mgmt_net_route_cidrs" :	"10.25.8.0/23" ,
		"int_backendic_gateway	10.20.0.1" : " " ,
		"int_backendic_net_name" :	"VPMS-FN-26071-T-BE-01_com_backend_ic_net_1" ,
		"int_backendic_route_cidrs" :	"10.20.0.1/24" ,
		"int_cdr_direct_gateway" :	"10.0.8.1" ,
		"int_cdr_direct_net_name" :	"VPMS-FN-26071-T-BE-01_com-cdr_direct_net_1" ,
		"int_cdr_direct_route_cidrs" :	"10.0.8.0/23" ,
		"int_vertcaic_gateway" :	"10.10.1.1" ,
		"int_vertcaic_net_name" :	"VPMS-FN-26071-T-BE-01_NBEA_verticaic_net_1" ,
		"int_vertcaic_route_cidrs" :	"10.10.1.0/24" ,
		"int_vertcaic_subnet_alloc_end" :	"10.10.1.254" ,
		"int_vertcaic_subnet_alloc_start" :	"10.10.1.3" ,
		"int_vertcaic_subnet_cidr" :	"10.10.1.0/24" ,
		"int_vertcaic_subnet_name" :	"VPMS-FN-26071-T-BE-01_NBEA_verticaic_net_1_subnet_1" ,
		"analyst_cluster_name" :	"zrdm5anbea01ana" ,
		"analyst_int_vprobes_mgmt_floating_ip" :	"10.25.8.11" ,
		"conductor_cluster_name" :	"zrdm5anbea01con" ,
		"conductor_int_vprobes_mgmt_floating_ip" :	"10.25.8.5" ,
		"rpmrepository_int_vprobes_mgmt_ip_0" :	"" ,
		"configurationrepository_cluster_name" :	"zrdm5anbea01crp" ,
		"distributedlock_cluster_name" :	"zrdm5anbea01dtl" ,
		"distributedlock_node_count" :	"3" ,
		"settingsandhealthdb_cluster_name" :	"zrdm5anbea01shd" ,
		"apigw_cluster_name" :	"zrdm5anbea01agw" ,
		"managementui_cluster_name" :	"zrdm5anbea01mgu" ,
		"timesten_cluster_name" :	"zrdm5anbea01ttn" ,
		"qtracelb_cluster_name" :	"zrdm5anbea01qlb" ,
		"guardian_cluster_name" :	"zrdm5anbea01gdn" ,
		"cognoscgw_cluster_name" :	"zrdm5anbea01cgw" ,
		"cognosccm_cluster_name" :	"zrdm5anbea01ccm" ,
		"cognoscdp_cluster_name" :	"zrdm5anbea01cdp" ,
		"vertica_cluster_name" :	"zrdm5anbea01vdb" ,
		"vertica_configuration_cluster_name" :	"zrdm5anbea01vdb" ,
		"vertica_data_cluster_name" :	"zrdm5anbea01vdb" ,
		"vertica_maintenance_cluster_name" :	"zrdm5anbea01vdb" ,
		"is_single_disk_app_dirs" :	"1" ,
		"vertica_multi_aff_index" :	"3" ,
		"vertica_volume_nr_1" :	"4" ,
		"initial_site_list" :	"[{'name':'RDM5a', 'role':'national', 'full_name':'RDM5a'},{'name':'default','role':'national', 'full_name':''}]" ,
		"is_fencing_enabled" :	"TRUE" ,
		"is_primary" :	"TRUE" ,
		"ntp_config_server" :	"155.179.59.249" ,
		"rsyslog_port" :	"514" ,
		"rsyslog_server" :	"None" ,
		"enable_cdr_reduction" :	"1" ,
		"enable_gatekeeper" :	"0" ,
		"enable_qalarm" :	"False" 
}

	rel_rows = pd.read_excel(path, sheet_name = "Tag-values", skiprows=6, nrows=26, usecols = 'B')
	b_params = rel_rows.iloc[:,0].tolist()
	# print(rel_rows)
	# print(b_params)
	# print('\n')

	loc = path
	wb = xlrd.open_workbook(loc)
	sheet_names = wb.sheet_names()
	#print('Sheet Names', sheet_names)
	tags_sheet = wb.sheet_by_name(sheet_names[8])
	#print(tags_sheet.cell_value(5,1)) # access B6 so it does row, column

	for key, value in tags_common_dict.items():
		for i in range(tags_sheet.nrows):
			if(tags_sheet.cell_value(i, 1)) == key and not None:
				#print(i, key, value)
				wb = xw.Book(path)
				wb.sheets[8].range('C'+ str(i+1)).value = value

def replace_tag_values_not_common(path):
	cluster_name_dict = {
		"rpmrepository" :	"zrdm5anbea01srp",
		"conductor" :	"zrdm5anbea01con",
		"settingsandhealthdb" :	"zrdm5anbea01shd",
		"apigw" :	"zrdm5anbea01agw",
		"analyst" :	"zrdm5anbea01ana",
		"configrepo" :	"zrdm5anbea01crp",
		"distributedlock" :	"zrdm5anbea01dtl",
		"vertica" :	"zrdm5anbea01vdb",
		"schemamanager" :	"zrdm5anbea01dbm",
		"microservices" :	"zrdm5anbea01msr",
		"managementui" :	"zrdm5anbea01mgu",
		"guardian" :	"zrdm5anbea01gdn",
		"qloader" :	"zrdm5anbea01ldr",
		"scheduledservices" :	"zrdm5anbea01ssr",
		"daemon" :	"zrdm5anbea01dmn",
		"timesten" :	"zrdm5anbea01ttn",
		"processing" :	"zrdm5anbea01mbm",
		"qtraceprocessing" :	"zrdm5anbea01qtp",
		"qtracelb" :	"zrdm5anbea01qlb",
		"cognoscgw" :	"zrdm5anbea01cgw",
		"cognosccm" :	"zrdm5anbea01ccm",
		"cognoscdp" :	"zrdm5anbea01cdp",
		"logfilter" :	"zrdm5anbea01log",
		"QTrace Server" :	"zrdm5anbea01qts",
		"Vertica Aux DB" :	"zrdm5anbea01adb",
		"Kafka" :	"zrdm5anbea01kfk",
		"Vertica MC" :	"zrdm5anbea01vmc",
		"SDN-O (Ansible)" :	 "zrdm5anbea01sdo",
		"SDN-O (Kubernete)" :	"zrdm5anbea01sdo",
		"Data Collection & Analysis" :	"zrdm5anbea01ljs",
		"Jumpserver Windows" :	"zrdm5anbea01wjs",
		"Jumpserver Linux" :	"zrdm5anbea01ljs"
	}

	flavor_dict = {
		"apigw" :	"nv.c8r32d80",
		"analyst" :	"nv.c8r32d80",
		"configrepo" :	"nv.c8r32d80",
		"distributedlock" :	"nv.c4r8d80",
		"vertica" :	"nv.c12r48d80 cpu pinning, NUMA",
		"schemamanager" :	 "nv.c8r32d80",
		"microservices" :	"nv.c2r4d80",
		"managementui" :	"nv.c8r32d80",
		"guardian" :	"nv.c8r32d80",
		"qloader" :	"nv.c8r32d80",
		"scheduledservices" : "nv.c4r8d80",
		"daemon" : "nv.c2r4d80",
		"timesten" : "nv.c8r32d80",
		"processing" :	"nv.c8r32d80e1500 cpu pinning, NUMA",
		"qtraceprocessing" : "nv.c8r32d80e1500",
		"qtracelb" : "nv.c4r8d80",
		"cognoscgw" :	"nv.c4r8d80",
		"cognosccm" :	"nv.c4r8d80",
		"cognoscdp" :	"nv.c8r32d80",
		"logfilter" :	"nv.c8r32d80e1500",
		"QTrace Server" :	"nv.c8r32d80e1500",
		"Vertica Aux DB" :	"nv.c8r32d80e1500",
		"Kafka" :	"nv.c8r16d200",
		"Vertica MC" :	"nv.c8r32d80e1500",
		"SDN-O (Ansible)" :	"nv.c8r16d60s1",
		"SDN-O (Kubernete)" :	"nv.c4r8d40s1",
		"Data Collection & Analysis" :	"nv.c8r32d80e1500",
		"Jumpserver Windows" :	"nv.c8r32d80e1500",
		"Jumpserver Linux" :	"nv.c8r32d80e1500"
		}

	images_dict = {
		"rpmrepository" :	"PROBE_SOFTWARE-REPOSITORY_11.3.7.000.02.qcow2",
		"conductor" :	"PROBE_VPROBE_11.3.7.qcow2",
		"settingsandhealthdb" :	"PROBE_VPROBE_11.3.7.qcow2",
		"apigw" :	"PROBE_VPROBE_11.3.7.qcow2",
		"analyst" :	"PROBE_VPROBE_11.3.7.qcow2",
		"configrepo" :	"PROBE_ORACLE_11.3.7.qcow2",
		"distributedlock" :	"PROBE_VPROBE_11.3.7.qcow2",
		"vertica" :	"PROBE_VPROBE_11.3.7_SCSI_DRIVER.qcow2",
		"schemamanager" :	"PROBE_VPROBE_11.3.7.qcow2",
		"microservices" :	"PROBE_VPROBE_11.3.7.qcow2",
		"managementui" :	"PROBE_VPROBE_11.3.7.qcow2",
		"guardian	" : "PROBE_VPROBE_11.3.7.qcow2",
		"qloader" :	"PROBE_VPROBE_11.3.7.qcow2",
		"scheduledservices" :	"PROBE_VPROBE_11.3.7.qcow2",
		"daemon" :	"PROBE_VPROBE_11.3.7.qcow2",
		"timesten" :	"PROBE_VPROBE_11.3.7.qcow2",
		"processing" :	"PROBE_VPROBE_11.3.7.qcow2",
		"qtraceprocessing" :	"PROBE_VPROBE_11.3.7.qcow2",
		"qtracelb" :	"PROBE_VPROBE_11.3.7.qcow2",
		"cognoscgw" :	"PROBE_VPROBE_11.3.7.qcow2",
		"cognosccm" :	"PROBE_VPROBE_11.3.7.qcow2",
		"cognoscdp" :	"PROBE_VPROBE_11.3.7.qcow2",
		"logfilter" :	"PROBE_VPROBE_11.3.7.qcow2",
		"QTrace Server" :	 "AT&T Windows Image",
		"Vertica Aux DB" :	"AT&T Linux Image",
		"Kafka" :	"AT&T Linux Image",
		"Vertica MC" :	"AT&T Linux Image",
		"SDN-O (Ansible)" : 	"AT&T Linux Image",
		"SDN-O (Kubernete)" :	"AT&T Linux Image",
		"Data Collection & Analysis" :	"AT&T Linux Image",
		"Jumpserver Windows" :	"AT&T Windows Image",
		"Jumpserver Linux" :	"AT&T Linux Image"
		}

	vm_name_dict = {
		"rpmrepository" : "zrdm5anbea01srp001" ,
		"conductor" :	"zrdm5anbea01con001,zrdm5anbea01con002" ,
		"settingsandhealthdb" :	"zrdm5anbea01shd001,zrdm5anbea01shd002,zrdm5anbea01shd003" ,
		"apigw" :	"zrdm5anbea01agw001,zrdm5anbea01agw002" ,
		"analyst" : "z0rdm5anbea01ana001,z0rdm5anbea01ana002" ,
		"configrepo" :	"zrdm5anbea01crp001" ,
		"distributedlock" : "zrdm5anbea01dtl001,zrdm5anbea01dtl002,zrdm5anbea01dtl003" , 
		"vertica" :	"zrdm5anbea01vdb001,zrdm5anbea01vdb002,zrdm5anbea01vdb003,zrdm5anbea01vdb004" ,
		"schemamanager" :	"zrdm5anbea01dbm001" ,
		"microservices" :	"zrdm5anbea01msr001,zrdm5anbea01msr002" ,
		"managementui" :	"zrdm5anbea01mgu001,zrdm5anbea01mgu002" ,
		"guardian" :	"zrdm5anbea01gdn001,zrdm5anbea01gdn002" ,
		"qloader" :	"zrdm5anbea01ldr001,zrdm5anbea01ldr002" ,
		"scheduledservices" :	"zrdm5anbea01ssr001, zrdm5anbea01ssr002" ,
		"daemon" :	"zrdm5anbea01dmn001,zrdm5anbea01dmn002" ,
		"timesten" :	 "zrdm5anbea01ttn001,zrdm5anbea01ttn002" ,
		"processing" :	"zrdm5anbea01mbm001,zrdm5anbea01mbm002" ,
		"qtraceprocessing" : "zrdm5anbea01qtp001,zrdm5anbea01qtp002" ,
		"qtracelb" :	"zrdm5anbea01qlb001,zrdm5anbea01qlb002" ,
		"cognoscgw" :	"zrdm5anbea01cgw001,zrdm5anbea01cgw002" ,
		"cognosccm" :	"zrdm5anbea01ccm001,zrdm5anbea01ccm002" ,
		"cognoscdp" :	"zrdm5anbea01cdp001,zrdm5anbea01cdp002" ,
		"logfilter" :	"zrdm5anbea01log001" ,
		"QTrace Server" :	"LLCCCCRAVPMW01,LLCCCCRAVPMW02" ,
		"Vertica Aux DB" :	"zrdm5anbea01adb001,zrdm5anbea01adb002,zrdm5anbea01adb003" ,
		"Kafka" :	"zrdm5anbea01kfk001,zrdm5anbea01kfk002,zrdm5anbea01kfk003" ,
		"Vertica MC" :	"zrdm5anbea01vmc001,zrdm5anbea01vmc002" ,
		"SDN-O (Ansible)" :	"zrdm5anbea01sda001,zrdm5anbea01sda002" ,
		"SDN-O (Kubernete)" :	"zrdm5anbea01sdk001,zrdm5anbea01sdk002,zrdm5anbea01sdk003" ,
		"Data Collection & Analysis" :	"zrdm5anbea01dca001,zrdm5anbea01dca002" ,
		"Jumpserver Windows" :	"LLCCCCCRAVPMA01,LLCCCCCRAVPMA02" ,
		"Jumpserver Linux" :	"zrdm5anbea01ljs001,zrdm5anbea01ljs002" 
	}

	protected_ips_dict = {
		"rpmrepository" :	"107.112.129.75",
		"conductor" :	"107.112.129.90,107.112.129.93",
		"apigw" :		"107.112.129.85,107.112.129.86",
		"analyst" :		"107.112.129.96,107.112.129.97",
		"managementui" : "107.112.129.99,107.112.129.104",
		"guardian" :	"107.112.129.",
		"processing" :	"107.112.129.76,107.112.129.77"	,
		"qtracelb" :	"107.112.129.78,107.112.129.79",
		"cognoscgw" : "107.112.129.82,107.112.129.83",
		"logfilter" :	"107.112.129.82"
	}

	floating_ips_dict = {
		"rpmrepository" :	"107.112.129.75",
		"conductor" :	"107.112.129.94",
		"apigw" :		"107.112.129.87",
		"analyst" :		"107.112.129.98",
		"managementui" : "107.112.129.108",
		"guardian" :	"107.112.129.129",
		#"processing" :	"107.112.129."	,
		"qtracelb" :	"107.112.129.81",
		"cognoscgw" : "107.112.129.84"
		#"logfilter" :	"107.112.129.82"
	}


	vm_count_dict = {
		"rpmrepository" :	"1",
		"conductor" :	"2",
		"settingsandhealthdb" :	"3",
		"apigw" :	"2",
		"analyst" :	"2",
		"configrepo" :	"1",
		"distributedlock" :	"3",
		"vertica" :	"4",
		"schemamanager" :	"1",
		"microservices" :	"2",
		"managementui" :	"2",
		"guardian" :	"2",
		"qloader" :	"2",
		"scheduledservices" :	"2",
		"daemon" :	"2",
		"timesten" :	"2",
		"processing" :	"2",
		"qtraceprocessing" :	"2",
		"qtracelb" :	"2",
		"cognoscgw" :	"2",
		"cognosccm" :	"2",
		"cognoscdp" :	"2",
		"logfilter" :	"1",
		"QTrace Server" :	"2",
		"Vertica Aux DB" :	"0",
		"Kafka" :	"0",
		"Vertica MC" :	"0",
		"SDN-O (Ansible)" :	"0",
		"SDN-O (Kubernete)" :	"0",
		"Data Collection & Analysis" :	"0",
		"Jumpserver Windows" :	"0",
		"Jumpserver Linux" :	"1"
	}



	extract_vm = pd.read_excel(path, sheet_name="VMs", usecols = 'B')
	col_B_list_vms = extract_vm.iloc[:, 0].tolist() # Save values of Column B to a list
	vm_type = col_B_list_vms[-1] # save VM Type
	#print(vm_type)

	loc = path
	wb = xlrd.open_workbook(loc)
	sheet_names = wb.sheet_names()
	tags_sheet = wb.sheet_by_name(sheet_names[8])
	vm_num_appendation = int(path[-6])
	vm_num_appendation_str = (path[-6])
	# REPLACE cluster name
	for key, value in cluster_name_dict.items():
		for i in range(tags_sheet.nrows): # iterate through all the rows so that order does not matter 
			if (tags_sheet.cell_value(i, 1)).endswith("cluster_name") and (tags_sheet.cell_value(i, 1)).startswith(vm_type) and key == vm_type: # checks start and end and vm type
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value
	# REPLACE flavor name
	for key, value in flavor_dict.items():
		for i in range(tags_sheet.nrows):	
			if (tags_sheet.cell_value(i, 1)).endswith("flavor_name") and (tags_sheet.cell_value(i, 1)).startswith(vm_type) and key == vm_type: # checks start and end and vm type
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value
	# REPLACE images name
	for key, value in images_dict.items():
		for i in range(tags_sheet.nrows):
			if (tags_sheet.cell_value(i, 1)).endswith("image_name") and (tags_sheet.cell_value(i, 1)).startswith(vm_type) and key == vm_type: # checks start and end and vm type
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value
	# REPLACE index name
	for i in range(tags_sheet.nrows):
		if (tags_sheet.cell_value(i, 1)).endswith("index") and (tags_sheet.cell_value(i, 1)).startswith(vm_type):
			wb = xw.Book(path)
			wb.sheets[8].range('C' + str(i+1)).value = (vm_num_appendation - 1)
	# REPLACE vm name
	for key, value in vm_name_dict.items():
		for i in range(tags_sheet.nrows):	
			if (tags_sheet.cell_value(i, 1)).endswith("names") and (tags_sheet.cell_value(i, 1)).startswith(vm_type) and key == vm_type:
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value # + vm_num_appendation_str
	# REPLACE vm count
	for key, value in vm_count_dict.items():
		for i in range(tags_sheet.nrows):	
			if (tags_sheet.cell_value(i, 1)).endswith("node_count") and (tags_sheet.cell_value(i, 1)).startswith(vm_type) and key == vm_type:
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value
	# REPLACE ips count
	for key, value in protected_ips_dict.items():
		for i in range(tags_sheet.nrows):	
			if (tags_sheet.cell_value(i, 1)).endswith("protected_ips") and (tags_sheet.cell_value(i, 1)).startswith("conductor_oam") and key == vm_type:
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value
	# REPLACE floated ips count
	for key, value in floating_ips_dict.items():
		for i in range(tags_sheet.nrows):	
			if (tags_sheet.cell_value(i, 1)).endswith("protected_floating_ip") and (tags_sheet.cell_value(i, 1)).startswith("conductor_oam") and key == vm_type:
				wb = xw.Book(path)
				wb.sheets[8].range('C' + str(i+1)).value = value

def utility_extract_value(path, sn, ci):
	wb = xw.Book(path) # open up a xlwings book, this is what gives you read and write permissions while preserving VBA
	print("Path:", path, "Value:", wb.sheets[sn].range(ci).value)



build_path =  r"C:\Users\rs623u\vnf_changes\files\buildplans.xlsx"
#path = r"C:\Users\rs623u\vnf_changes\files\tt_zrdm5anbea01_daemon_1.xlsm" #copy_zrdm5anbea01_vertica_1.xlsm"
#path = r"C:\Users\rs623u\vnf_changes\vProbe_WAH1a_FE__EPDG_preload_11_3_7\vProbe_FE3_Addon_vProbe_4_module_WAH1a_FE__EPDG_preload_11_3_7.xlsm"


path = r"C:\Users\rs623u\vnf_changes\vProbe_WAH1a_FE__EPDG_preload_11_3_7\fe3"
files = []
print("Please Enter Sheet # (0 indexed)")
sn = int(input())
print("Please Enter Cell Index (ex. C25)")
ci = str(input())
for idx, item in enumerate(os.listdir(path)):
	files.append(item)
	fullpath = os.path.join(path, item)
	utility_extract_value(fullpath, sn, ci)
	#print(fullpath)
# 	os.rename(fullpath, fullpath.replace(item, name_list[idx]))

# extract_common_general(build_path)
# replace_module_model_general(path)
# replace_module_name_general(path)
# replace_networks(path)
# change_vm_name(path)
# change_az(path)
# change_ips(path)
# replace_tag_values(path)
# replace_tag_values_not_common(path)
# for i in paths:
# 	extract_value_c25( "r" + i )
