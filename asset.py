import wmi
import json

from win32com.client import GetObject

objWMI = GetObject('winmgmts:\\\\.\\root\\SecurityCenter2').InstancesOf('AntiVirusProduct')

c = wmi.WMI()

cpu = c.Win32_Processor()[0].Name

displayname = ''
instanceguid = ''
productState = ''
timestamp = ''

for obj in objWMI:
	if obj.displayName != None:
		displayname = str(obj.displayName)
	if obj.instanceGuid != None:
		instanceguid = str(obj.instanceGuid)
	if obj.productState != None:
		productState = str(obj.productState)
	if obj.timestamp != None:
		timestamp = str(obj.timestamp)
	print("")
	print("########")
	print("")

data = [{"CPU": cpu, "AV": displayname, "State": productState,}]

with open('test.json', 'w') as f:
	json.dump(data, f, indent=4)

print("CPU: " + cpu)
print("Anti Virus: " + displayname)
print("Product State: " + productState)
print("Time Stamp: " + timestamp)

input("Press Enter to continue...")