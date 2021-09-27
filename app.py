from flask import Flask, render_template, redirect, request, session,flash
import requests
import json
import os
import pandas as pd
import xlrd
from flask import request
import zipfile
import shutil
from flask import send_file
import numpy as np
from flask import * 
from werkzeug.utils import secure_filename
import openpyxl
from flask_session import Session
import time




app = Flask(__name__)
t1=0
app.config['UPLOAD_EXTENSIONS'] = ['.xlsx']
app.secret_key = 'SECRET KEY'
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)
   
def createcommand(command, key, value):
        
    if (value!='dummy'):
        
        
        if isinstance(value,list):
            value_tmp=''
            for i in value:
                
                value_tmp=str(value_tmp + " --" + key + ' ' + str(i))
            return str(command + value_tmp)
        elif isinstance(value,float):
            
            value = str(int(value))
            return str(command + " --" + key + ' ' + value )
        else:
            value=str(value)
            return str(command + " --" + key + ' ' + value )
    else:
        
        return command


@app.route("/")
@app.route("/home")
def home():
    return render_template("home.html")

@app.route("/generate", methods = ["GET", "POST"])
def generate():
    
    global t1
        
    if not session.get("file"):
    
         return render_template("home.html",mg="No File Found ! Upload File !")    
        
    if request.method == "POST":
        inputFile=session.get("file")
        sheet = request.form.getlist('sheet')
        if not sheet:
            return render_template("home.html",mg="No Checkbox Selected !") 
            
    
        
        
        

#getting sheet names
        path = session['path']+"/"
        xls = xlrd.open_workbook(inputFile, on_demand=True)
        sheet_names = xls.sheet_names()

           
        zipf = zipfile.ZipFile('efacli.zip','w', zipfile.ZIP_DEFLATED)
        df_global = pd.read_excel(inputFile, sheet_name='Global', engine='openpyxl')
        df_global = df_global.replace(np.nan, 'dummy')
        
        temp_po = {}
        for xx in range(0, len(df_global['po-number'])):
            if df_global['po-number'][xx] != 'dummy':
                temp_po[int(df_global['po-number'][xx])] = df_global['po-name'][xx]
                
        temp_mask = {}   
        for yy in range(0, len(df_global['vlan-id'])):
            if df_global['subnet'][yy] != 'dummy':
                temp_mask[str(df_global['vlan-id'][yy])] = df_global['subnet'][yy].split('/')[1]
                
        temp_mask_v6 = {} 
        for zz in range(0, len(df_global['vlan-id'])):
            if df_global['subnet-ipv6'][zz] != 'dummy':
                temp_mask_v6[str(df_global['vlan-id'][zz])] = df_global['subnet-ipv6'][zz].split('/')[1]
                                                                     
        
        
        #create a new excel file for every sheet
        for name1 in sheet_names:
            xlrd.xlsx.ensure_elementtree_imported(False, None)
            xlrd.xlsx.Element_has_iter = True
            
            if name1 in sheet :
                
    
                

                #writing data to the new excel file
                ## From here code will change it will be specific to a sheet
                if name1 == "Tenant":#########################################
                    
                    
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df1 = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    col_list=list(df1.columns)
                    df = df.replace(np.nan, 'dummy')
                    
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        
                        tenant = []
                        for k in range(0,len(df['efa tenant create'])):
                            
                            
    
                            tenant_temp = {}
                            tenant_temp['name']=df['name'][k]
                            tenant_temp['description']=df['description'][k]
                            tenant_temp['type']=df['type'][k]
                            tenant_temp['l2-vni-range']=df['l2-vni-range'][k]
                            tenant_temp['l3-vni-range']=df['l3-vni-range'][k]
                            tenant_temp['vlan-range']=df['vlan-range'][k]
                            tenant_temp['vrf-count']=df['vrf-count'][k]
                            tenant_temp['enable-bd']=(str(df['enable-bd'][k])).lower()
                            tenant_temp['port_sw']=df['sw-ip'][k]+"["+df['sw-port'][k]+"]"
                            tenant.append(tenant_temp)
	                       
	                       
                        
                        lst1 = []
                        for i in tenant:
                            
                            if (i['name'] !="dummy"):
                                
                                command = "efa tenant create"
                                name = createcommand(command,'name',i['name'])
                                description = createcommand(name,'description',i['description'])
                                type = createcommand(description,'type',i['type'])
                                l2_vni_range = createcommand(type,'l2-vni-range',i['l2-vni-range'])
                                l3_vni_range = createcommand(l2_vni_range,'l3-vni-range',i['l3-vni-range'])
                                vlan_range = createcommand(l3_vni_range,'vlan-range',i['vlan-range'])
                                vrf_count = createcommand(vlan_range,'vrf-count',i['vrf-count'])
                                enable_bd = createcommand(vrf_count,'enable-bd',i['enable-bd'])
                                port = createcommand(enable_bd,'port',i['port_sw'])
                                lst1.append(port)
        
                            else:
                                
                                port = "," + i['port_sw']
                                last_element = lst1[-1]
                                lst1.pop()
                                tmp = last_element + port
                                lst1.append(tmp)
                                
                        for i in lst1:
                            outfile.write("%s\n\n" % i)

                            
                            
				             
                            
                    zipf.write(path+str(name1)+".txt")
                    print(request)
                    ##Here the code will end for a specific sheet type
                if name1 == "Inventory":
                    df2 = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df2 = df2.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        inventory = []
                        for k in range(0, len(df2['efa inventory device register'])):
                            inventory_temp = {}
                            inventory_temp['ip'] = df2['sw-ip'][k]
                            inventory_temp['username'] = df2['username'][k]
                            inventory_temp['password'] = df2['password'][k]
                            inventory.append(inventory_temp)
                        lst7 = []
                        for i in inventory:
                            if i['ip']!= 'dummy':
                                command = "efa inventory device register"
                                ip = createcommand(command,'ip',i['ip'])
                                username = createcommand(ip,'username',i['username'])
                                password = createcommand(username,'password',i['password'])
                                lst7.append(password)
                    
                        for i in lst7:
                            outfile.write("%s\n\n" % i)
                            
                    
                    zipf.write(path+str(name1)+".txt")
                    print(request)
                    
                if name1 == "Breakout":
                    df3 = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df3 = df3.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        breakout = []
                        for k in range(0, len(df3['efa inventory device interface'])):
                            breakout_temp = {}
                            breakout_temp['sw-ip'] = " --ip " + df3['sw-ip'][k] + " --if-type eth --if-name " 
                            breakout_temp['sw-ip-br'] = " --ip " + df3['sw-ip'][k]
                            breakout_temp['4x25g'] = df3['4x25g'][k]
                            breakout_temp['4x10g'] = df3['4x10g'][k]
                            breakout_temp['4x1g'] = df3['4x1g'][k]
                            breakout.append(breakout_temp)
                        
                        lst8 = []
                        for i in breakout:
                            command = "efa inventory device interface"
                            if i['4x25g'] != "dummy":
                                x25g_state_down = command + " set-admin-state" + i['sw-ip'] + i['4x25g'] + " --state down"
                                x25g_breakout_mode = x25g_state_down + "\n" + command + " set-breakout" + i['sw-ip'] + i['4x25g'] + " --mode 4x25g"
                                x25g_state_up = x25g_breakout_mode + "\n" + command + " set-admin-state" + i['sw-ip'] + i['4x25g'] + ":1-4 --state up" + "\n"
                                lst8.append(x25g_state_up)
                    
                            if i['4x10g'] != "dummy":
                                x10g_state_down = command + " set-admin-state" + i['sw-ip'] + i['4x10g'] + " --state down"
                                x10g_breakout_mode = x10g_state_down + "\n" + command + " set-breakout" + i['sw-ip'] + i['4x10g'] + " --mode 4x10g"
                                x10g_state_up = x10g_breakout_mode + "\n" + command + " set-admin-state" + i['sw-ip'] + i['4x10g'] + ":1-4 --state up"  + "\n"
                                lst8.append(x10g_state_up)
                                
                            if i['4x1g'] != "dummy":
                                x1g_state_down = command + " set-admin-state" + i['sw-ip'] + i['4x1g'] + " --state down"
                                x1g_breakout_mode = x1g_state_down + "\n" + command + " set-breakout" + i['sw-ip'] + i['4x1g'] + " --mode 4x1g"
                                x1g_red_mgmt = x1g_breakout_mode + "\n" + "efa inventory device execute-cli" + i['sw-ip-br'] + ' --command "config, interface Ethernet ' + i['4x1g'] + ':1, description Redundant_Mgmt_Interface, redundant-management enable"'
                                x1g_state_up = x1g_red_mgmt + "\n" + command + " set-admin-state" + i['sw-ip'] + i['4x1g'] + ":1-4 --state up" + "\n"
                                lst8.append(x1g_state_up)
                        
                        update_fabric = "efa inventory device update --fabric " + df_global['fabric-name'][0]
                        lst8.append(update_fabric)
                        
                        for i in lst8:
                            outfile.write("%s\n\n" % i)
                            
                    zipf.write(path+str(name1)+".txt")
                    print(request)   
                    
                    
                if name1 == "Fabric":
                    df4 = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df4 = df4.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        if df_global['stage'][0] != 'dummy':
                            fabric_create = "efa fabric create --name " + df_global['fabric-name'][0] + " --type " + df_global['fabric-type'][0] + " --stage " + str(int(df_global['stage'][0])) + "\n"
                        else:
                            fabric_create = "efa fabric create --name " + df_global['fabric-name'][0] + " --type " + df_global['fabric-type'][0] + "\n"
                        outfile.write(fabric_create)  
                        
                        fabric = {}
                        for i in range(0, len(df4['flags'])):
                            if df4['value'][i] != 'dummy':
                                fabric[df4['value'][i]] = df4['flags'][i]
                                
                        lst9 = ["efa fabric setting update"]
                        for i in fabric:
                            if isinstance(i,float):
                                val = str(int(i))
                                fabric_update = " --" + fabric[i] + " " + val
                            else:
                                val = str(i)
                                fabric_update = " --" + fabric[i] + " " + val
                            lst9.append(fabric_update)    
                        
                        device = []
                        for k in range(0, len(df4['device-ip'])):
                            if df4['device-ip'][k] != "dummy":
                                device_temp = {}
                                device_temp['device-ip'] = df4['device-ip'][k]
                                device_temp['device-role'] = df4['device-role'][k]
                                device_temp['leaf-type'] = df4['leaf-type'][k]
                                device_temp['hostname'] = df4['hostname'][k]
                                device_temp['asn'] = df4['asn'][k]
                                device_temp['vtep-loopback'] = df4['vtep-loopback'][k]
                                device_temp['loopback'] = df4['loopback'][k]
                                device_temp['pod'] = df4['pod'][k]
                                device_temp['username'] = df4['username'][k]
                                device_temp['password'] = df4['password'][k]
                                device_temp['rack'] = df4['rack'][k]
                                device.append(device_temp)
                                
                        device_lst = []
                        for i in device:
                            command = "efa fabric device add --name " + df4['fabric-name'][0]
                            device_ip = createcommand(command,'ip',i['device-ip'])
                            device_role = createcommand(device_ip,'role',i['device-role'])
                            leaf_type = createcommand(device_role,'leaf-type',i['leaf-type'])
                            hostname = createcommand(leaf_type,'hostname',i['hostname'])
                            asn = createcommand(hostname,'asn',i['asn'])
                            vtep_loopback = createcommand(asn,'vtep-loopback',i['vtep-loopback'])
                            loopback = createcommand(vtep_loopback,'loopback',i['loopback'])
                            pod = createcommand(loopback,'pod',i['pod'])
                            username = createcommand(pod,'username',i['username'])
                            password = createcommand(username,'password',i['password'])
                            rack = createcommand(password,'rack',i['rack'])
                            device_lst.append(rack)
                            
                        for i in lst9:
                            outfile.write("%s" % i)
                        for j in device_lst:
                            outfile.write("\n%s" % j)
                        efa_configure = "\n"+ "efa fabric configure --name " + df_global['fabric-name'][0]
                        outfile.write(efa_configure)
                        
                    zipf.write(path+str(name1)+".txt")
                    print(request)   
                    
                
                if name1 == "Tenant PO":
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df = df.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        tenant_po = []
                        for k in range(0, len(df['efa tenant po create'])):
                            tenant_po_temp = {}
                            tenant_po_temp['tenant'] = df['tenant'][k]
                            tenant_po_temp['name'] = df['name'][k]
                            tenant_po_temp['description'] = df['description'][k]
                            tenant_po_temp['speed'] = df['speed'][k]
                            tenant_po_temp['negotiation'] = df['negotiation'][k]
                            
                            if df['sw2-ip'][k] != 'dummy': tenant_po_temp['port-sw'] = df['sw1-ip'][k] + "[" + df['sw1-port'][k] + "]" +  "," + df['sw2-ip'][k] + "[" + df['sw2-port'][k] + "]"
                            else: tenant_po_temp['port-sw'] = df['sw1-ip'][k] + "[" + df['sw1-port'][k] + "]"
                            
                            tenant_po_temp['min-link-count'] = df['min-link-count'][k]
                            tenant_po_temp['number'] = df['number'][k]
                            tenant_po_temp['lacp-timeout'] = df['lacp-timeout'][k]
                            tenant_po.append(tenant_po_temp)
                            
                        lst3 = []    
                        for i in tenant_po:
                            command = "efa tenant po create"
                            tenant = createcommand(command,'tenant',i['tenant'])
                            name = createcommand(tenant,'name',i['name'])
                            description = createcommand(name,'description',i['description'])
                            speed = createcommand(description,'speed',i['speed'])
                            negotiation = createcommand(speed,'negotiation',i['negotiation'])
                            port = createcommand(negotiation,'port',i['port-sw'])
                            min_link_count = createcommand(port,'min-link-count',i['min-link-count'])
                            number = createcommand(min_link_count,'number',i['number'])
                            lacp_timeout = createcommand(number,'lacp-timeout',i['lacp-timeout'])
                            lst3.append(lacp_timeout)
                            
                        for i in lst3:
                            outfile.write("%s\n" % i)
                            
                    zipf.write(path+str(name1)+".txt")
                    print(request)
                    
                    
                if name1 == "Tenant BGP Peer-Group":
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df = df.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        tenant_bgp_pg = []
                        for k in range(0, len(df['efa tenant service bgp peer-group create'])):
                            tenant_bgp_pg_temp = {}
                            tenant_bgp_pg_temp['name'] = df['name'][k]
                            tenant_bgp_pg_temp['tenant'] = df['tenant'][k]
                            tenant_bgp_pg_temp['description'] = df['description'][k]
                            if   (df['pg-name'][k] != 'dummy') : tenant_bgp_pg_temp['pg-name'] = df['device-ip'][k] + ":" + df['pg-name'][k]
                            else : tenant_bgp_pg_temp['pg-name'] = 'dummy'
                            if   (df['pg-asn'][k] != 'dummy') : tenant_bgp_pg_temp['pg-asn'] = df['device-ip'][k] + "," + df['pg-name'][k] + ":" + str(df['pg-asn'][k])
                            else : tenant_bgp_pg_temp['pg-asn'] = 'dummy'
                            if   (df['pg-bfd'][k] != 'dummy') : tenant_bgp_pg_temp['pg-bfd'] = df['device-ip'][k] + "," + df['pg-name'][k] + ":" + df['pg-bfd'][k]
                            else : tenant_bgp_pg_temp['pg-bfd'] = 'dummy'
                            if   (df['pg-bfd-enable'][k] != 'dummy') : tenant_bgp_pg_temp['pg-bfd-enable'] = df['device-ip'][k] + "," + df['pg-name'][k] + ":" + str(df['pg-bfd-enable'][k]).lower()
                            else : tenant_bgp_pg_temp['pg-bfd-enable'] = 'dummy'
                            if   (df['pg-next-hop-self'][k] != 'dummy') : tenant_bgp_pg_temp['pg-next-hop-self'] = df['device-ip'][k] + "," + df['pg-name'][k] + ":" + str(df['pg-next-hop-self'][k]).lower()
                            else : tenant_bgp_pg_temp['pg-next-hop-self'] = 'dummy'
                            if   (df['pg-update-source-ip'][k] != 'dummy') : tenant_bgp_pg_temp['pg-update-source-ip'] = df['device-ip'][k] + "," + df['pg-name'][k] + ":" + df['pg-update-source-ip'][k]
                            else : tenant_bgp_pg_temp['pg-update-source-ip'] = 'dummy'
                            
                            tenant_bgp_pg.append(tenant_bgp_pg_temp)
                            
                        lst5 = []
                        for i in tenant_bgp_pg:
                            if i['tenant'] != "dummy":
                                command = "efa tenant service bgp peer-group create"
                                tenant = createcommand(command,'tenant',i['tenant'])
                                name = createcommand(tenant,'name',i['name'])
                                description = createcommand(name,'description',i['description'])
                                pg_name = createcommand(description,'pg-name',i['pg-name'])
                                pg_asn = createcommand(pg_name,'pg-asn',i['pg-asn'])
                                pg_bfd = createcommand(pg_asn,'pg-bfd',i['pg-bfd'])
                                pg_bfd_enable = createcommand(pg_bfd,'pg-bfd-enable',i['pg-bfd-enable'])
                                pg_nhs = createcommand(pg_bfd_enable,'pg-next-hop-self',i['pg-next-hop-self'])
                                pg_update_source_ip_sw = createcommand(pg_nhs,'pg-update-source-ip',i['pg-update-source-ip'])
                                lst5.append(pg_update_source_ip_sw)

                            else:
                                pg_name = createcommand("",'pg-name',i['pg-name'])
                                pg_asn = createcommand(pg_name,'pg-asn',i['pg-asn'])
                                pg_bfd = createcommand(pg_asn,'pg-bfd',i['pg-bfd'])
                                pg_bfd_enable = createcommand(pg_bfd,'pg-bfd-enable',i['pg-bfd-enable'])
                                pg_nhs = createcommand(pg_bfd_enable,'pg-next-hop-self',i['pg-next-hop-self'])
                                pg_update_source_ip_sw = createcommand(pg_nhs,'pg-update-source-ip',i['pg-update-source-ip'])

                                last_element = lst5[-1]
                                lst5.pop()
                                tmp = last_element + pg_update_source_ip_sw
                                lst5.append(tmp)
                                
                        for i in lst5:
                            outfile.write("%s\n\n" % i)
                                
                    zipf.write(path+str(name1)+".txt")
                    print(request)        

                    
                if name1 == "Tenant EPG":
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl')
                    df = df.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        tenant_epg = []
                        for k in range(0, len(df['efa tenant epg create'])):
                            tenant_epg_temp = {}
                            tenant_epg_temp['name'] = df['name'][k]
                            tenant_epg_temp['tenant'] = df['tenant'][k]
                            tenant_epg_temp['description'] = df['description'][k]
                            if   ((df['sw1-port'][k] != 'dummy') and (df['sw2-port'][k] != 'dummy')) : tenant_epg_temp['port-sw'] = df['sw1-ip'][k] + "[" + df['sw1-port'][k] + "]" +  "," + df['sw2-ip'][k] + "[" + df['sw2-port'][k] + "]"
                            elif ((df['sw1-port'][k] != 'dummy') and (df['sw2-port'][k] == 'dummy')) : tenant_epg_temp['port-sw'] = df['sw1-ip'][k] + "[" + df['sw1-port'][k] + "]"
                            elif ((df['sw1-port'][k] == 'dummy') and (df['sw2-port'][k] != 'dummy')) : tenant_epg_temp['port-sw'] = df['sw2-ip'][k] + "[" + df['sw2-port'][k] + "]"			
                            elif ((df['sw1-port'][k] == 'dummy') and (df['sw2-port'][k] == 'dummy')) : tenant_epg_temp['port-sw'] = 'dummy'
                    
                            temp_var = df['po-number'][k]
                            po_lst = []
                            if "," in temp_var:
                                j = 0
                                for i in temp_var.split(','):
                                    j+=1
                                    po_lst.append(temp_po[int(i)])
                                tenant_epg_temp['po'] = po_lst
                            elif temp_var !='dummy' and "," not in temp_var:
                                tenant_epg_temp['po'] = temp_po[int(temp_var)]
                            else:
                                tenant_epg_temp['po'] = df['po-number'][k]
                             
                                
                            tenant_epg_temp['switchport-mode'] = df['switchport-mode'][k]
                            tenant_epg_temp['type'] = df['type'][k]
                            tenant_epg_temp['switchport-native-vlan-tagging'] = df['switchport-native-vlan-tagging'][k]
                            tenant_epg_temp['single-homed-bfd-session-type'] = df['single-homed-bfd-session-type'][k]
                            tenant_epg_temp['switchport-native-vlan'] = df['switchport-native-vlan'][k]
                            ctag_lst = []
                            anycast_lst = []
                            anycast_ipv6_lst = []
                            sw1_local_lst = []
                            sw2_local_lst = []
                            sw1_local_ipv6_lst = []
                            sw2_local_ipv6_lst = []
                            
                            if '-' in str(df['ctag-range'][k]):
                                start = str(df['ctag-range'][k]).split('-')[0]
                                end = str(df['ctag-range'][k]).split('-')[1]
                                for i in range(int(start),int(end)+1):
                                    ctag_lst.append(str(i))
                                tenant_epg_temp['ctag-range'] = df['ctag-range'][k] 
                            else:
                                
                                tenant_epg_temp['ctag-range'] = (str(df['ctag-range'][k]))
                            
                            if '[' in str(df['anycast-ip'][k]):
                                start_ctag = int(str(df['ctag-range'][k]).split('-')[0])
                                start = str(df['anycast-ip'][k]).split('.[')[1].split('-')[0].strip()
                                end = str(df['anycast-ip'][k]).split('].')[0].split('-')[1].strip()
                                for i in range(int(start),int(end)+1):  
                                    temp_var = str(start_ctag) + ":" + str(df['anycast-ip'][k]).split('.[')[0] + "." + str(i) + "." + str(df['anycast-ip'][k]).split('].')[1] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                                    anycast_lst.append(temp_var)
                                    start_ctag +=1
                                tenant_epg_temp['anycast-ip'] = anycast_lst
                            else:
                                if df['anycast-ip'][k] == "dummy":
                                    tenant_epg_temp['anycast-ip'] = "dummy"
                                else:
                                    tenant_epg_temp['anycast-ip'] = str(df['ctag-range'][k]) + ":" + df['anycast-ip'][k] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                            
                            if '[' in str(df['anycast-ipv6'][k]):  
                                start_ctag = int(str(df['ctag-range'][k]).split('-')[0])
                                start = str(df['anycast-ipv6'][k]).split(':[')[1].split('-')[0].strip()
                                end = str(df['anycast-ipv6'][k]).split(']:')[0].split('-')[1].strip()
                                for i in range(int(start),int(end)+1):
                                    temp_var = str(start_ctag) + ":" + str(df['anycast-ipv6'][k]).split(':[')[0] + ":" + str(i) + ":" + str(df['anycast-ipv6'][k]).split(']:')[1] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                                    anycast_lst.append(temp_var)
                                    start_ctag +=1
                                tenant_epg_temp['anycast-ipv6'] = anycast_lst
                            else:
                                if df['anycast-ipv6'][k] == "dummy":
                                    tenant_epg_temp['anycast-ipv6'] = "dummy"
                                else:
                                    tenant_epg_temp['anycast-ipv6'] = str(df['ctag-range'][k]) + ":" + df['anycast-ipv6'][k] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                                  
                            if '[' in str(df['sw1-local-ip'][k]):
                                start_ctag = int(str(df['ctag-range'][k]).split('-')[0])
                                start = str(df['sw1-local-ip'][k]).split('.[')[1].split('-')[0].strip()
                                end = str(df['sw1-local-ip'][k]).split('].')[0].split('-')[1].strip()
                                for i in range(int(start),int(end)+1):
                                    temp_var = str(start_ctag) + "," + df['sw1-ip'][k] + ":" + str(df['sw1-local-ip'][k]).split('.[')[0] + "." + str(i) + "." + str(df['sw1-local-ip'][k]).split('].')[1] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                                    sw1_local_lst.append(temp_var)
                                    start_ctag +=1
                                tenant_epg_temp['sw1-local-ip'] = sw1_local_lst
                            else: 
                                if df['sw1-local-ip'][k] == "dummy":
                                    tenant_epg_temp['sw1-local-ip'] = "dummy"
                                else:
                                    tenant_epg_temp['sw1-local-ip'] = str(df['ctag-range'][k]) + "," + df['sw1-ip'][k] + ":" + df['sw1-local-ip'][k] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                            
                            if '[' in str(df['sw2-local-ip'][k]):
                                start_ctag = int(str(df['ctag-range'][k]).split('-')[0])
                                start = str(df['sw2-local-ip'][k]).split('.[')[1].split('-')[0].strip()
                                end = str(df['sw2-local-ip'][k]).split('].')[0].split('-')[1].strip()
                                for i in range(int(start),int(end)+1):
                                    temp_var = str(start_ctag) + "," + df['sw2-ip'][k] + ":" + str(df['sw2-local-ip'][k]).split('.[')[0] + "." + str(i) + "." + str(df['sw2-local-ip'][k]).split('].')[1] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                                    sw2_local_lst.append(temp_var)
                                    start_ctag +=1
                                tenant_epg_temp['sw2-local-ip'] = sw2_local_lst
                            
                            else:   
                                if df['sw2-local-ip'][k] == "dummy":
                                    tenant_epg_temp['sw2-local-ip'] = "dummy"
                                else:
                                    tenant_epg_temp['sw2-local-ip'] = str(df['ctag-range'][k]) + "," + df['sw2-ip'][k] + ":" + df['sw2-local-ip'][k] + "/" + str(temp_mask[str(df['ctag-range'][k])])
                            
                            if '[' in str(df['sw1-local-ipv6'][k]):  
                                start_ctag = int(str(df['ctag-range'][k]).split('-')[0])
                                start = str(df['sw1-local-ipv6'][k]).split(':[')[1].split('-')[0].strip()
                                end = str(df['sw1-local-ipv6'][k]).split(']:')[0].split('-')[1].strip()
                                for i in range(int(start),int(end)+1):  
                                    temp_var = str(start_ctag) + "," + df['sw1-ip'][k] + ":" + str(df['sw1-local-ipv6'][k]).split(':[')[0] + ":" + str(i) + ":" + str(df['sw1-local-ipv6'][k]).split(']:')[1] + "/" + str(temp_mask_v6[str(df['ctag-range'][k])])
                                    sw1_local_ipv6_lst.append(temp_var)
                                    start_ctag +=1
                                tenant_epg_temp['sw1-local-ipv6'] = sw1_local_ipv6_lst
                            else:
                                if df['sw1-local-ipv6'][k] == "dummy":
                                    tenant_epg_temp['sw1-local-ipv6'] = "dummy"
                                else:
                                    tenant_epg_temp['sw1-local-ipv6'] = str(df['ctag-range'][k]) + "," + df['sw1-ip'][k] + ":" + df['sw1-local-ipv6'][k] + "/" + str(temp_mask_v6[str(df['ctag-range'][k])])
                                    
                            if '[' in str(df['sw2-local-ipv6'][k]): 
                                start_ctag = int(str(df['ctag-range'][k]).split('-')[0])
                                start = str(df['sw2-local-ipv6'][k]).split(':[')[1].split('-')[0].strip()
                                end = str(df['sw2-local-ipv6'][k]).split(']:')[0].split('-')[1].strip()
                                for i in range(int(start),int(end)+1):  
                                    temp_var = str(start_ctag) + "," + df['sw2-ip'][k] + ":" + str(df['sw2-local-ipv6'][k]).split(':[')[0] + ":" + str(i) + ":" + str(df['sw2-local-ipv6'][k]).split(']:')[1] + "/" + str(temp_mask_v6[str(df['ctag-range'][k])])
                                    sw2_local_ipv6_lst.append(temp_var)
                                    start_ctag +=1
                                tenant_epg_temp['sw2-local-ipv6'] = sw2_local_ipv6_lst
                            else:
                                if df['sw2-local-ipv6'][k] == "dummy":
                                    tenant_epg_temp['sw2-local-ipv6'] = "dummy"
                                else:
                                    tenant_epg_temp['sw2-local-ipv6'] = str(df['ctag-range'][k]) + "," + df['sw2-ip'][k] + ":" + df['sw2-local-ipv6'][k] + "/" + str(temp_mask_v6[str(df['ctag-range'][k])])
                            
                            tenant_epg_temp['vrf'] = df['vrf'][k]
                            tenant_epg_temp['l3-vni'] = df['l3-vni'][k]
                            tenant_epg_temp['l2-vni'] = df['l2-vni'][k]
                            
                            if df['ip-mtu'][k] == "dummy":
                                tenant_epg_temp['ip-mtu'] = "dummy"
                            else:
                                tenant_epg_temp['ip-mtu'] = str(df['ctag-range'][k]) + ":" +  str(int(df['ip-mtu'][k]))
                            
                            tenant_epg_temp['bridge-domain'] = df['bridge-domain'][k]
                            tenant_epg_temp['ipv6-nd-mtu'] = df['ipv6-nd-mtu'][k]
                            tenant_epg_temp['ipv6-nd-managed-config'] = df['ipv6-nd-managed-config'][k]
                            tenant_epg_temp['ipv6-nd-other-config'] = df['ipv6-nd-other-config'][k]
                            tenant_epg_temp['ipv6-nd-prefix'] = df['ipv6-nd-prefix'][k]
                            tenant_epg_temp['ipv6-nd-prefix-valid-lifetime'] = df['ipv6-nd-prefix-valid-lifetime'][k]
                            tenant_epg_temp['ipv6-nd-prefix-preferred-lifetime'] = df['ipv6-nd-prefix-preferred-lifetime'][k]
                            tenant_epg_temp['ipv6-nd-prefix-no-advertise'] = df['ipv6-nd-prefix-no-advertise'][k]
                            tenant_epg_temp['ipv6-nd-prefix-config-type'] = df['ipv6-nd-prefix-config-type'][k]
                            tenant_epg_temp['ctag-description'] = df['ctag-description'][k]
                            tenant_epg.append(tenant_epg_temp)
                        
                        
                        lst4 = []    
                        for i in tenant_epg:
                            command = "efa tenant epg create"
                            tenant = createcommand(command,'tenant',i['tenant'])
                            name = createcommand(tenant,'name',i['name'])
                            description = createcommand(name,'description',i['description'])
                            port = createcommand(description,'port',i['port-sw'])
                            if (isinstance(i['po'],list)):
                                po = createcommand(port,'po',(",".join(i['po'])))
                            else:
                                po = createcommand(port,'po',(i['po']))
                            
                            switchport_mode = createcommand(po,'switchport-mode',i['switchport-mode'])
                            type = createcommand(switchport_mode,'type',i['type'])
                            switchport_tagging = createcommand(type,'switchport-native-vlan-tagging',i['switchport-native-vlan-tagging'])
                            single_homed_bfd = createcommand(switchport_tagging,'single-homed-bfd-session-type',i['single-homed-bfd-session-type'])
                            switchport_vlan = createcommand(single_homed_bfd,'switchport-native-vlan',i['switchport-native-vlan'])
                            ctag_range = createcommand(switchport_vlan,'ctag-range',i['ctag-range'])
                            vrf = createcommand(ctag_range,'vrf',i['vrf'])
                            l3_vni = createcommand(vrf,'l3-vni',i['l3-vni'])
                            l2_vni = createcommand(l3_vni,'l2-vni',i['l2-vni'])
                            ip_mtu = createcommand(l2_vni,'ip-mtu',i['ip-mtu'])
                            anycast_ip = createcommand(ip_mtu,'anycast-ip',i['anycast-ip'])
                            anycast_ipv6 = createcommand(anycast_ip,'anycast-ipv6',i['anycast-ipv6'])
                            sw1_local_ip = createcommand(anycast_ipv6,'local-ip',i['sw1-local-ip'])
                            sw2_local_ip = createcommand(sw1_local_ip,'local-ip',i['sw2-local-ip'])
                            sw1_local_ipv6 = createcommand(sw2_local_ip,'local-ipv6',i['sw1-local-ipv6'])
                            sw2_local_ipv6 = createcommand(sw1_local_ipv6,'local-ipv6',i['sw2-local-ipv6'])
                            bridge_domain = createcommand(sw2_local_ipv6,'bridge-domain',i['bridge-domain'])
                            ipv6_nd_mtu = createcommand(bridge_domain,'ipv6-nd-mtu',i['ipv6-nd-mtu'])
                            ipv6_nd_managed_config = createcommand(ipv6_nd_mtu,'ipv6-nd-managed-config',i['ipv6-nd-managed-config'])
                            ipv6_nd_other_config = createcommand(ipv6_nd_managed_config,'ipv6-nd-other-config',i['ipv6-nd-other-config'])
                            ipv6_nd_prefix = createcommand(ipv6_nd_other_config,'ipv6-nd-prefix',i['ipv6-nd-prefix'])
                            ipv6_nd_prefix_valid_lifetime = createcommand(ipv6_nd_prefix,'ipv6-nd-prefix-valid-lifetime',i['ipv6-nd-prefix-valid-lifetime'])
                            ipv6_nd_prefix_preferred_lifetime = createcommand(ipv6_nd_prefix_valid_lifetime,'ipv6-nd-prefix-preferred-lifetime',i['ipv6-nd-prefix-preferred-lifetime'])
                            ipv6_nd_prefix_no_advertise = createcommand(ipv6_nd_prefix_preferred_lifetime,'ipv6-nd-prefix-no-advertise',i['ipv6-nd-prefix-no-advertise'])
                            ipv6_nd_prefix_config_type = createcommand(ipv6_nd_prefix_no_advertise,'ipv6-nd-prefix-config-type',i['ipv6-nd-prefix-config-type'])
                            ctag_description = createcommand(ipv6_nd_prefix_config_type,'ctag-description',i['ctag-description'])
                            lst4.append(ctag_description)
                            
                            
                            
                        for i in lst4:
                            
                            outfile.write("%s\n\n" % i)
                    
                    zipf.write(path+str(name1)+".txt")
                    print(request)        
                            
                    
                if name1 == "Tenant VRF":
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl',header=[0, 1])
                    df = df.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        tenant_vrf = []
                        for k in range(0, len(df['tenant-vrf']['efa tenant vrf create'])):
                            tenant_vrf_temp = {}
                            tenant_vrf_temp['name'] = df['tenant-vrf']['name'][k]
                            tenant_vrf_temp['tenant'] = df['tenant-vrf']['tenant'][k]
                            tenant_vrf_temp['local-asn'] = df['tenant-vrf']['local-asn'][k]
                            tenant_vrf_temp['rt-type-import'] = df['tenant-vrf']['rt-type-import'][k]
                            tenant_vrf_temp['rt-import'] = df['tenant-vrf']['rt-import'][k]
                            tenant_vrf_temp['rt-type-export'] = df['tenant-vrf']['rt-type-export'][k]
                            tenant_vrf_temp['rt-export'] = df['tenant-vrf']['rt-export'][k]
                            tenant_vrf_temp['rt-type-both'] = df['tenant-vrf']['rt-type-both'][k]
                            tenant_vrf_temp['rt-both'] = df['tenant-vrf']['rt-both'][k]
                            tenant_vrf_temp['rh-max-path'] = df['tenant-vrf']['rh-max-path'][k]
                            tenant_vrf_temp['rh-ecmp-enable'] = str(df['tenant-vrf']['rh-ecmp-enable'][k]).lower()
                            tenant_vrf_temp['graceful-restart-enable'] = str(df['tenant-vrf']['graceful-restart-enable'][k]).lower()
                            tenant_vrf_temp['max-path'] = df['tenant-vrf']['max-path'][k]
                            tenant_vrf_temp['redistribute'] = df['tenant-vrf']['redistribute'][k]
                            tenant_vrf_temp['routing-type'] = df['tenant-vrf']['routing-type'][k]
                            tenant_vrf_temp['centralized-router'] = df['tenant-vrf']['centralized-router'][k]
                            
                            if   ((df['tenant-vrf']['device-ip'][k] != 'dummy') and (df['ipv4-static-route-next-hop']['static-route'][k] != 'dummy') and (df['ipv4-static-route-next-hop']['next-hop'][k] != 'dummy') and (df['ipv4-static-route-next-hop']['distance'][k] != 'dummy')) : tenant_vrf_temp['ipv4_srnh'] = df['tenant-vrf']['device-ip'][k] + "," + df['ipv4-static-route-next-hop']['static-route'][k] + "," + df['ipv4-static-route-next-hop']['next-hop'][k] +  "," + df['ipv4-static-route-next-hop']['distance'][k] 
                            elif ((df['tenant-vrf']['device-ip'][k] != 'dummy') and (df['ipv4-static-route-next-hop']['static-route'][k] != 'dummy') and (df['ipv4-static-route-next-hop']['next-hop'][k] != 'dummy') and (df['ipv4-static-route-next-hop']['distance'][k] == 'dummy')) : tenant_vrf_temp['ipv4_srnh'] = df['tenant-vrf']['device-ip'][k] + "," + df['ipv4-static-route-next-hop']['static-route'][k] + "," + df['ipv4-static-route-next-hop']['next-hop'][k]
                            else : tenant_vrf_temp['ipv4_srnh'] = 'dummy'
                            
                            
                            if   ((df['tenant-vrf']['device-ip'][k] != 'dummy') and (df['ipv6-static-route-next-hop']['static-route'][k] != 'dummy') and (df['ipv6-static-route-next-hop']['next-hop'][k] != 'dummy') and (df['ipv6-static-route-next-hop']['distance'][k] != 'dummy')) : tenant_vrf_temp['ipv6_srnh'] = df['tenant-vrf']['device-ip'][k] + "," + df['ipv6-static-route-next-hop']['static-route'][k] + "," + df['ipv6-static-route-next-hop']['next-hop'][k] +  "," + df['ipv6-static-route-next-hop']['distance'][k] 
                            elif ((df['tenant-vrf']['device-ip'][k] != 'dummy') and (df['ipv6-static-route-next-hop']['static-route'][k] != 'dummy') and (df['ipv6-static-route-next-hop']['next-hop'][k] != 'dummy') and (df['ipv6-static-route-next-hop']['distance'][k] == 'dummy')) : tenant_vrf_temp['ipv6_srnh'] = df['tenant-vrf']['device-ip'][k] + "," + df['ipv6-static-route-next-hop']['static-route'][k] + "," + df['ipv6-static-route-next-hop']['next-hop'][k]
                            else : tenant_vrf_temp['ipv6_srnh'] = 'dummy'
                            
                            
                            if   ((df['tenant-vrf']['device-ip'][k] != 'dummy') and (df['ipv4-static-route-bfd']['dest-addr'][k] != 'dummy') and (df['ipv4-static-route-bfd']['source-addr'][k] != 'dummy') and (df['ipv4-static-route-bfd']['interval,min-rx,multiplier'][k] != 'dummy')) : tenant_vrf_temp['ipv4_srb'] = df['tenant-vrf']['device-ip'][k] + "," + df['ipv4-static-route-bfd']['dest-addr'][k] + "," + df['ipv4-static-route-bfd']['source-addr'][k] +  "," + df['ipv4-static-route-bfd']['interval,min-rx,multiplier'][k]
                            else : tenant_vrf_temp['ipv4_srb'] = 'dummy'
                            
                            if   ((df['tenant-vrf']['device-ip'][k] != 'dummy') and (df['ipv6-static-route-bfd']['dest-addr'][k] != 'dummy') and (df['ipv6-static-route-bfd']['source-addr'][k] != 'dummy') and (df['ipv6-static-route-bfd']['interval,min-rx,multiplier'][k] != 'dummy')) : tenant_vrf_temp['ipv6_srb'] = df['tenant-vrf']['device-ip'][k] + "," + df['ipv6-static-route-bfd']['dest-addr'][k] + "," + df['ipv6-static-route-bfd']['source-addr'][k] +  "," + df['ipv6-static-route-bfd']['interval,min-rx,multiplier'][k]
                            else : tenant_vrf_temp['ipv6_srb'] = 'dummy'
                            
                            tenant_vrf.append(tenant_vrf_temp)
                            
                            
                            
                        lst2 = []
                        for i in tenant_vrf:
                            if i['tenant'] != "dummy":
                                command = "efa tenant vrf create"
                                tenant = createcommand(command,'tenant',i['tenant'])
                                name = createcommand(tenant,'name',i['name'])
                                local_asn = createcommand(name,'local-asn',i['local-asn'])
                                rt_type_imp = createcommand(local_asn,'rt-type',i['rt-type-import'])
                                rt_imp = createcommand(rt_type_imp,'rt',i['rt-import'])
                                rt_type_exp = createcommand(rt_imp,'rt-type',i['rt-type-export'])
                                rt_exp = createcommand(rt_type_exp,'rt',i['rt-export'])
                                rt_type_both = createcommand(rt_exp,'rt-type',i['rt-type-both'])
                                rt_both = createcommand(rt_type_both,'rt',i['rt-both'])
                                rh_mp = createcommand(rt_both,'rh-max-path',i['rh-max-path'])
                                rh_ecmp = createcommand(rh_mp,'rh-ecmp-enable', i['rh-ecmp-enable'])
                                graceful_re = createcommand(rh_ecmp,'graceful-restart-enable', i['graceful-restart-enable'])
                                max_path = createcommand(graceful_re,'max-path',i['max-path'])
                                redistribute = createcommand(max_path,'redistribute',i['redistribute'])
                                routing_type = createcommand(redistribute,'routing-type',i['routing-type'])
                                centralized_router = createcommand(routing_type,'centralized-router',i['centralized-router'])
                                ipv4_srnh = createcommand(centralized_router,'ipv4-static-route-next-hop',i['ipv4_srnh'])
                                ipv6_srnh = createcommand(ipv4_srnh,'ipv6-static-route-next-hop',i['ipv6_srnh'])
                                ipv4_srb = createcommand(ipv6_srnh,'ipv4-static-route-bfd',i['ipv4_srb'])
                                ipv6_srb = createcommand(ipv4_srb,'ipv6-static-route-bfd',i['ipv6_srb'])
                        
                                lst2.append(ipv6_srb)
                                
                            else:
                                redistribute = createcommand("",'redistribute',i['redistribute'])
                                ipv4_srnh = createcommand(redistribute,'ipv4-static-route-next-hop',i['ipv4_srnh'])
                                ipv6_srnh = createcommand(ipv4_srnh,'ipv6-static-route-next-hop',i['ipv6_srnh'])
                                ipv4_srb = createcommand(ipv6_srnh,'ipv4-static-route-bfd',i['ipv4_srb'])
                                ipv6_srb = createcommand(ipv4_srb,'ipv6-static-route-bfd',i['ipv6_srb'])
                                
                                last_element = lst2[-1]
                                lst2.pop()
                                tmp = last_element + ipv6_srb
                                lst2.append(tmp)
                                
                                
                        for i in lst2:
                            outfile.write("%s\n\n" % i)
                            
                            
                    zipf.write(path+str(name1)+".txt")
                    print(request)        
                
                    
                if name1 == "Tenant BGP Peer":
                    df = pd.read_excel(inputFile, sheet_name=name1, engine='openpyxl',header=[0, 1])
                    df = df.replace(np.nan, 'dummy')
                    with open(path+str(name1)+".txt", 'w') as outfile:
                        tenant_bgp_peer = []
                        for k in range(0, len(df['service-bgp']['efa tenant service bgp peer create'])):
                            tenant_bgp_peer_temp = {}
                            tenant_bgp_peer_temp['name'] = df['service-bgp']['name'][k]
                            tenant_bgp_peer_temp['tenant'] = df['service-bgp']['tenant'][k]
                            tenant_bgp_peer_temp['description'] = df['service-bgp']['description'][k]
                            
                            if   ((df['ipv4-uc-dyn-nbr']['ipv4-listen-range'][k] != 'dummy') and (df['ipv4-uc-dyn-nbr']['peer-group-name'][k] != 'dummy') and (df['ipv4-uc-dyn-nbr']['listen-limit'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-dyn-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-dyn-nbr']['ipv4-listen-range'][k] + "," + df['ipv4-uc-dyn-nbr']['peer-group-name'][k] + "," + str(int(df['ipv4-uc-dyn-nbr']['listen-limit'][k]))
                            elif ((df['ipv4-uc-dyn-nbr']['ipv4-listen-range'][k] != 'dummy') and (df['ipv4-uc-dyn-nbr']['peer-group-name'][k] != 'dummy') and (df['ipv4-uc-dyn-nbr']['listen-limit'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-dyn-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-dyn-nbr']['ipv4-listen-range'][k] + "," + df['ipv4-uc-dyn-nbr']['peer-group-name'][k]
                            else : tenant_bgp_peer_temp['ipv4-uc-dyn-nbr'] = 'dummy'
			
                            if   ((df['ipv4-uc-nbr']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr']['remote-as'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr']['ipv4-neighbor'][k] + "," + str(int(df['ipv4-uc-nbr']['remote-as'][k]))
                            elif ((df['ipv4-uc-nbr']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr']['remote-as'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr']['ipv4-neighbor'][k] 
                            else : tenant_bgp_peer_temp['ipv4-uc-nbr'] = 'dummy' 

                            if   ((df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-bfd']['bfd-enable'][k] != 'dummy') and (df['ipv4-uc-nbr-bfd']['interval,min-rx,multiplier'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] + "," + str(df['ipv4-uc-nbr-bfd']['bfd-enable'][k]).lower() + "," + df['ipv4-uc-nbr-bfd']['interval,min-rx,multiplier'][k]
                            elif ((df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-bfd']['bfd-enable'][k] != 'dummy') and (df['ipv4-uc-nbr-bfd']['interval,min-rx,multiplier'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] + "," + str(df['ipv4-uc-nbr-bfd']['bfd-enable'][k]).lower() 
                            elif ((df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-bfd']['bfd-enable'][k] == 'dummy') and (df['ipv4-uc-nbr-bfd']['interval,min-rx,multiplier'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] + "," + df['ipv4-uc-nbr-bfd']['interval,min-rx,multiplier'][k] 
                            elif ((df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-bfd']['bfd-enable'][k] == 'dummy') and (df['ipv4-uc-nbr-bfd']['interval,min-rx,multiplier'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-bfd']['ipv4-neighbor'][k] 
                            else : tenant_bgp_peer_temp['ipv4-uc-nbr-bfd'] = 'dummy'

                            if   ((df['ipv4-uc-nbr-next-hop-self']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-next-hop-self']['next-hop-self'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-next-hop-self'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-next-hop-self']['ipv4-neighbor'][k] + "," + str(df['ipv4-uc-nbr-next-hop-self']['next-hop-self'][k]).lower() 
                            elif ((df['ipv4-uc-nbr-next-hop-self']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-next-hop-self']['next-hop-self'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-next-hop-self'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-next-hop-self']['ipv4-neighbor'][k]
                            else : tenant_bgp_peer_temp['ipv4-uc-nbr-next-hop-self'] = 'dummy' 

                            if   ((df['ipv4-uc-nbr-update-source-ip']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-update-source-ip']['update-source-ip'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-update-source-ip'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-update-source-ip']['ipv4-neighbor'][k] + "," + df['ipv4-uc-nbr-update-source-ip']['update-source-ip'][k] 				
                            elif ((df['ipv4-uc-nbr-update-source-ip']['ipv4-neighbor'][k] != 'dummy') and (df['ipv4-uc-nbr-update-source-ip']['update-source-ip'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv4-uc-nbr-update-source-ip'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv4-uc-nbr-update-source-ip']['ipv4-neighbor'][k]
                            else : tenant_bgp_peer_temp['ipv4-uc-nbr-update-source-ip'] = 'dummy' 

                            if   ((df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] != 'dummy') and (df['ipv6-uc-dyn-nbr']['peer-group-name'][k] != 'dummy') and (df['ipv6-uc-dyn-nbr']['listen-limit'][k] != 'dummy')) :	tenant_bgp_peer_temp['ipv6-uc-dyn-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] + "," + df['ipv6-uc-dyn-nbr']['peer-group-name'][k] + "," + str(int(df['ipv6-uc-dyn-nbr']['listen-limit'][k]))
                            elif ((df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] != 'dummy') and (df['ipv6-uc-dyn-nbr']['peer-group-name'][k] != 'dummy') and (df['ipv6-uc-dyn-nbr']['listen-limit'][k] == 'dummy')) :	tenant_bgp_peer_temp['ipv6-uc-dyn-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] + "," + df['ipv6-uc-dyn-nbr']['peer-group-name'][k] 
                            elif ((df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] != 'dummy') and (df['ipv6-uc-dyn-nbr']['peer-group-name'][k] == 'dummy') and (df['ipv6-uc-dyn-nbr']['listen-limit'][k] != 'dummy')) :	tenant_bgp_peer_temp['ipv6-uc-dyn-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] + "," + df['ipv6-uc-dyn-nbr']['listen-limit'][k] 
                            elif ((df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k] != 'dummy') and (df['ipv6-uc-dyn-nbr']['peer-group-name'][k] == 'dummy') and (df['ipv6-uc-dyn-nbr']['listen-limit'][k] == 'dummy')) :	tenant_bgp_peer_temp['ipv6-uc-dyn-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-dyn-nbr']['ipv6-listen-range'][k]
                            else : tenant_bgp_peer_temp['ipv6-uc-dyn-nbr'] = 'dummy'

                            if   ((df['ipv6-uc-nbr']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr']['remote-as'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr']['ipv6-neighbor'][k] + "," + str(int(df['ipv6-uc-nbr']['remote-as'][k]))
                            elif ((df['ipv6-uc-nbr']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr']['remote-as'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr']['ipv6-neighbor'][k] 
                            else : tenant_bgp_peer_temp['ipv6-uc-nbr'] = 'dummy' 

                            if   ((df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-bfd']['bfd-enable'][k] != 'dummy') and (df['ipv6-uc-nbr-bfd']['interval,min-rx,multiplier'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] + "," + str(df['ipv6-uc-nbr-bfd']['bfd-enable'][k]).lower() + "," + df['ipv6-uc-nbr-bfd']['interval,min-rx,multiplier'][k]
                            elif ((df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-bfd']['bfd-enable'][k] != 'dummy') and (df['ipv6-uc-nbr-bfd']['interval,min-rx,multiplier'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] + "," + str(df['ipv6-uc-nbr-bfd']['bfd-enable'][k]).lower() 
                            elif ((df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-bfd']['bfd-enable'][k] == 'dummy') and (df['ipv6-uc-nbr-bfd']['interval,min-rx,multiplier'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] + "," + df['ipv6-uc-nbr-bfd']['interval,min-rx,multiplier'][k] 
                            elif ((df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-bfd']['bfd-enable'][k] == 'dummy') and (df['ipv6-uc-nbr-bfd']['interval,min-rx,multiplier'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-bfd'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-bfd']['ipv6-neighbor'][k] 
                            else : tenant_bgp_peer_temp['ipv6-uc-nbr-bfd'] = 'dummy'

                            if   ((df['ipv6-uc-nbr-next-hop-self']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-next-hop-self']['next-hop-self'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-next-hop-self'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-next-hop-self']['ipv6-neighbor'][k] + "," + str(df['ipv6-uc-nbr-next-hop-self']['next-hop-self'][k]).lower()
                            elif ((df['ipv6-uc-nbr-next-hop-self']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-next-hop-self']['next-hop-self'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-next-hop-self'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-next-hop-self']['ipv6-neighbor'][k]
                            else : tenant_bgp_peer_temp['ipv6-uc-nbr-next-hop-self'] = 'dummy' 

                            if   ((df['ipv6-uc-nbr-update-source-ip']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-update-source-ip']['update-source-ip'][k] != 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-update-source-ip'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-update-source-ip']['ipv6-neighbor'][k] + "," + df['ipv6-uc-nbr-update-source-ip']['update-source-ip'][k] 				
                            elif ((df['ipv6-uc-nbr-update-source-ip']['ipv6-neighbor'][k] != 'dummy') and (df['ipv6-uc-nbr-update-source-ip']['update-source-ip'][k] == 'dummy')) : tenant_bgp_peer_temp['ipv6-uc-nbr-update-source-ip'] = df['service-bgp']['device-ip'][k] + "," + df['service-bgp']['vrf-name'][k] + ":" + df['ipv6-uc-nbr-update-source-ip']['ipv6-neighbor'][k]
                            else : tenant_bgp_peer_temp['ipv6-uc-nbr-update-source-ip'] = 'dummy'

                            tenant_bgp_peer.append(tenant_bgp_peer_temp)
                            
                            
                        lst6 = []
                        for i in tenant_bgp_peer:
                            if i['tenant'] != "dummy":
                                command = "efa tenant service bgp peer create"
                                tenant = createcommand(command,'tenant',i['tenant'])
                                name = createcommand(tenant,'name',i['name'])
                                description = createcommand(name,'description',i['description'])
                                ipv4_uc_dyn_nbr = createcommand(description,'ipv4-uc-dyn-nbr',i['ipv4-uc-dyn-nbr'])
                                ipv4_uc_nbr = createcommand(ipv4_uc_dyn_nbr,'ipv4-uc-nbr',i['ipv4-uc-nbr'])
                                ipv4_uc_nbr_bfd = createcommand(ipv4_uc_nbr,'ipv4-uc-nbr-bfd',i['ipv4-uc-nbr-bfd'])
                                ipv4_uc_nbr_nhs = createcommand(ipv4_uc_nbr_bfd,'ipv4-uc-nbr-next-hop-self',i['ipv4-uc-nbr-next-hop-self'])
                                ipv4_uc_nbr_update_source_ip = createcommand(ipv4_uc_nbr_nhs,'ipv4-uc-nbr-update-source-ip',i['ipv4-uc-nbr-update-source-ip'])
                                ipv6_uc_dyn_nbr = createcommand(ipv4_uc_nbr_update_source_ip,'ipv6-uc-dyn-nbr',i['ipv6-uc-dyn-nbr'])
                                ipv6_uc_nbr = createcommand(ipv6_uc_dyn_nbr,'ipv6-uc-nbr',i['ipv6-uc-nbr'])
                                ipv6_uc_nbr_bfd = createcommand(ipv6_uc_nbr,'ipv6-uc-nbr-bfd',i['ipv6-uc-nbr-bfd'])
                                ipv6_uc_nbr_nhs = createcommand(ipv6_uc_nbr_bfd,'ipv6-uc-nbr-next-hop-self',i['ipv6-uc-nbr-next-hop-self'])
                                ipv6_uc_nbr_update_source_ip = createcommand(ipv6_uc_nbr_nhs,'ipv6-uc-nbr-update-source-ip',i['ipv6-uc-nbr-update-source-ip'])
                                lst6.append(ipv6_uc_nbr_update_source_ip)
                                
                                
                            else:
                                ipv4_uc_dyn_nbr = createcommand("",'ipv4-uc-dyn-nbr',i['ipv4-uc-dyn-nbr'])
                                ipv4_uc_nbr = createcommand(ipv4_uc_dyn_nbr,'ipv4-uc-nbr',i['ipv4-uc-nbr'])
                                ipv4_uc_nbr_bfd = createcommand(ipv4_uc_nbr,'ipv4-uc-nbr-bfd',i['ipv4-uc-nbr-bfd'])
                                ipv4_uc_nbr_nhs = createcommand(ipv4_uc_nbr_bfd,'ipv4-uc-nbr-next-hop-self',i['ipv4-uc-nbr-next-hop-self'])
                                ipv4_uc_nbr_update_source_ip = createcommand(ipv4_uc_nbr_nhs,'ipv4-uc-nbr-update-source-ip',i['ipv4-uc-nbr-update-source-ip'])
                                ipv6_uc_dyn_nbr = createcommand(ipv4_uc_nbr_update_source_ip,'ipv6-uc-dyn-nbr',i['ipv6-uc-dyn-nbr'])
                                ipv6_uc_nbr = createcommand(ipv6_uc_dyn_nbr,'ipv6-uc-nbr',i['ipv6-uc-nbr'])
                                ipv6_uc_nbr_bfd = createcommand(ipv6_uc_nbr,'ipv6-uc-nbr-bfd',i['ipv6-uc-nbr-bfd'])
                                ipv6_uc_nbr_nhs = createcommand(ipv6_uc_nbr_bfd,'ipv6-uc-nbr-next-hop-self',i['ipv6-uc-nbr-next-hop-self'])
                                ipv6_uc_nbr_update_source_ip = createcommand(ipv6_uc_nbr_nhs,'ipv6-uc-nbr-update-source-ip',i['ipv6-uc-nbr-update-source-ip'])

                                last_element = lst6[-1]
                                lst6.pop()
                                tmp = last_element + ipv6_uc_nbr_update_source_ip
                                lst6.append(tmp)
                                
                                
                                
                        for i in lst6:
                            outfile.write("%s\n\n" % i)
                            
                            
                    zipf.write(path+str(name1)+".txt")
                    print(request)        
                    
    zipf.close()    
    shutil.rmtree(session['path'])
    session.pop('path',None)
    session.pop('file',None)
    return send_file('efacli.zip',
            mimetype = 'zip',
            attachment_filename= 'efacli.zip',
            as_attachment = True)


@app.route('/download')
def download_file():
    p="_EFA.xlsx"
    return send_file(p,as_attachment=True)
    

@app.route('/uploader', methods = ['GET', 'POST'])   
def upload_file():
    
    global t1
     
    if request.method == 'POST':
       file = request.files['file']
       filename1=secure_filename(file.filename)
       if filename1 != '':
           file_ext=os.path.splitext(filename1)[1]
           if file_ext not in app.config['UPLOAD_EXTENSIONS']:
               return render_template("home.html", messaage="Wrong File type, only 'xlsx' format allowed !")
           
       if not file.filename:
           return render_template("home.html", messaage="Select File First !")     
       t1=t1+1
           
       if os.path.isdir(str(t1)):
           for i in range(0,400):
               present = time.time()
               olderThanDays=900
               subDirPath="%d/"%t1
               if (present - os.path.getmtime(subDirPath)) > olderThanDays:
                   shutil.rmtree(subDirPath)
                   
               t1=t1+1
               if os.path.isdir(str(t1)):
                   continue 
               break
              
       os.mkdir(str(t1))
       session['path']="%d"%t1
       session["file"]="%d/%s"%(t1,file.filename)
       if not session.get("file"):
           return render_template("home.html", messaage="Select File First !")
          
                  
       file.save("%d/%s"%(t1,file.filename))
          
       if session.get("file"):
           file1=session["file"]
           flash(f"{file1} uploaded successfully, now you can generate sheets","info")
           return redirect(url_for("home"))        
    else:
       
        return render_template("home.html")




if __name__ == '__main__':
    os.environ['FLASK_ENV'] = 'development'
    app.run(debug=True)
    