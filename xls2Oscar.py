import sys
import os
import xlrd
import yaml
from functools import reduce  # forward compatibility for Python 3
import operator
import pygeometa
from datetime import datetime

####### open xls2yml #######
file=sys.argv[1]
os.system("python xls2yml.py "+file)

############ run pygeometa ##########

from pygeometa.core import read_mcf, render_j2_template
# read from disk
mcf_dict = read_mcf(os.path.splitext(file)[0]+'.yml')
# choose wmo-wigos output schema
from pygeometa.schemas.wmo_wigos import WMOWIGOSOutputSchema
iso_os = WMOWIGOSOutputSchema()
# default schema
xml_string = iso_os.write(mcf_dict)
# write to disk
with open(os.path.splitext(file)[0]+'.xml', 'w') as ff:
    ff.write(xml_string)


##### load into Oscar using pyOscar #####
from pyoscar import OSCARClient
client = OSCARClient(api_token='foo', env='depl')
#client = OSCARClient(api_token='foo', env='prod') ### uncomment this line for production environment ######
with open(os.path.splitext(file)[0]+".xml") as fh:
    data=fh.read()
response=client.upload(data)


