import sys
import os
import xlrd
import yaml
from functools import reduce  # forward compatibility for Python 3
import operator
import pygeometa
from datetime import datetime

####### open xls file listed in argument #######
file=sys.argv[1]
xl_workbook = xlrd.open_workbook(file)
sheet_names = xl_workbook.sheet_names()
new_yaml = open(os.path.splitext(file)[0]+'pre.yml', 'w')

def nested_set(dic, keys, value):
    for key in keys[:-1]:
        dic = dic.setdefault(key, {})
    dic[keys[-1]] = value

my_dic={}

#### HEADER ####
nested_set(my_dic,["mcf","version"],"1.0")
nested_set(my_dic,["metadata","identifier"],"3f342f64-9348-11df-ba6a-0014c2c00eab")
nested_set(my_dic,["metadata","language"],"en")
nested_set(my_dic,["metadata","language_alternate"],"fr")
nested_set(my_dic,["metadata","charset"],"utf8")
now=datetime.now()
dtstring=now.strftime("%Y-%m-%dT%H:%M:%S%Z")
nested_set(my_dic,["metadata","datestamp"],dtstring)

### set up contacts
xl_sheet = xl_workbook.sheet_by_name("contact")
for x in range(2,xl_sheet.nrows):
    if xl_sheet.cell(x,0).value=="":
        continue
    for y in range(xl_sheet.ncols):
        cell=xl_sheet.cell(x,y).value
        if cell!="":
            parents=["contact","main",xl_sheet.cell(0,y).value,xl_sheet.cell(1,y).value]
            parents = list(filter(None, parents))
            nested_set(my_dic, parents,xl_sheet.cell(x,y).value)

            parents=["contact","facility",xl_sheet.cell(0,y).value,xl_sheet.cell(1,y).value]
            parents = list(filter(None, parents))
            nested_set(my_dic, parents,xl_sheet.cell(x,y).value)

            parents=["contact","record_owner",xl_sheet.cell(0,y).value,xl_sheet.cell(1,y).value]
            parents = list(filter(None, parents))
            nested_set(my_dic, parents,xl_sheet.cell(x,y).value)

namessofar=[]
#open facility sheet
xl_sheet = xl_workbook.sheet_by_name("facility")
name="facility"
for x in range(7,xl_sheet.nrows):
    if xl_sheet.cell(x,0).value=="":
        continue
    stationname=xl_sheet.cell(x,1).value
    for y in range(xl_sheet.ncols):
        cell=xl_sheet.cell(x,y).value
        if cell!="":
            parents=[name,xl_sheet.cell(x,1).value,xl_sheet.cell(0,y).value,xl_sheet.cell(1,y).value,xl_sheet.cell(2,y).value,xl_sheet.cell(3,y).value,xl_sheet.cell(4,y).value]

            #eliminate empty
            parents = list(filter(None, parents))

            #Check if station already is in the list, add ' for each timeperiod/validperiod to add
            if stationname in namessofar:
                if "- timeperiod" in parents:
                    parents = [w.replace("- timeperiod","- timeperiod"+"'"*namessofar.count(stationname)) for w in parents]
                if "- valid_period" in parents:
                    parents = [w.replace("- valid_period","- valid_period"+"'"*namessofar.count(stationname)) for w in parents]
            nested_set(my_dic, parents,xl_sheet.cell(x,y).value)


#open variables sheet
xl_sheet = xl_workbook.sheet_by_name("observations")
name="observations"
for x in range(7,xl_sheet.nrows):
    if xl_sheet.cell(x,0).value=="":
        continue
    stationname=xl_sheet.cell(x,0).value
    for y in range(xl_sheet.ncols):
        cell=xl_sheet.cell(x,y).value
        if cell!="":
            parents=["facility",stationname,xl_sheet.cell(0,y).value,xl_sheet.cell(1,y).value,xl_sheet.cell(2,y).value,xl_sheet.cell(3,y).value,xl_sheet.cell(4,y).value]

            #eliminate empty
            parents = list(filter(None, parents))

            #Check if station already is in the list, add ' for each timeperiod/validperiod to add
            if stationname in namessofar:
                if "- timeperiod" in parents:
                    parents = [w.replace("- timeperiod","- timeperiod"+"'"*namessofar.count(stationname)) for w in parents]
                if "- valid_period" in parents:
                    parents = [w.replace("- valid_period","- valid_period"+"'"*namessofar.count(stationname)) for w in parents]
            nested_set(my_dic, parents,xl_sheet.cell(x,y).value)


yaml.dump(my_dic, new_yaml)
new_yaml.close()

#take away the quotation marks
with open(os.path.splitext(file)[0]+'pre.yml', 'r') as f, open(os.path.splitext(file)[0]+'.yml', 'w') as fo:
    for line in f:
        fo.write(line.replace('"', '').replace("'", ""))
