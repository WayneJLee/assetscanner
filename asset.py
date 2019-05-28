import wmi
import sys
import json
import re
# import csv

from pprint import pprint as pp
from win32com.client import GetObject

c = wmi.WMI()

objWMI = GetObject('winmgmts:\\\\.\\root\\SecurityCenter2').InstancesOf('AntiVirusProduct')

wql = "SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'"

hostname = c.Win32_ComputerSystem()[0].Caption
os = c.Win32_OperatingSystem()[0].Caption
cpu = c.Win32_Processor()[0].Name
ram = int(int(c.Win32_ComputerSystem()[0].TotalPhysicalMemory) / 1000000)
ip = [ip.IPAddress for ip in c.query(wql)]
installedapps = c.Win32_Product()
displayname = ''
avdetected = ''
locationfound = False

# Removing non-string name fields to allow regex iteration
for index, app in enumerate(installedapps):
    if app.Name == None:
        installedapps.pop(index)

# Detecting Anti Virus presence and name
for obj in objWMI:
    if obj.displayName != None:
        displayname = str(obj.displayName)
        avdetected = True

# Search for installed location of Anti Virus
# if displayname != 'Windows Defender' and avdetected == True:
for app in installedapps:
    searchav = displayname + '.*'
    p = re.compile(searchav)
    m = p.match(app.Name)
    if m:

        location = app.InstallLocation
        locationfound = True

# Assigning collected data to variable

data = [{
    'HostName': hostname,
    'OperatingSystem': os,
    'CPU': cpu,
    'RAM': ram,
    'IPAddress': ip,
    'AntiVirus': displayname,
}]

if locationfound == True:
    data[0].update({'Location': location})


with open(hostname + '.json', 'w') as f:
    json.dump(data, f, indent=4)

# CSV OUTPUT
# with open(hostname + '.csv', 'w') as w:
#     fields = ['HostName', 'OperatingSystem', 'RAM', 'IPAddress', 'AntiVirus', 'Version', 'InstallationDate', 'Location']
#     writer = csv.DictWriter(w, fieldnames=fields)
#     writer.writeheader()
#     writer.writerows(data)

# w.close()

pp(data)
