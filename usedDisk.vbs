' Run this script at the command prompt by typing
'     cscript usedDisk.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select FreeSpace,Size from Win32_LogicalDisk where Name = 'C:'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "FreeSpace" then
			freeC = oProperty.Value
		elseif oProperty.Name = "Size" then
			sizeC = oProperty.Value
		end if
	next
	percentUsedC = 100 - (100 * (freeC/sizeC))
next

wqlQuery = "select FreeSpace,Size from Win32_LogicalDisk where Name = 'E:'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "FreeSpace" then
			freeE = oProperty.Value
		elseif oProperty.Name = "Size" then
			sizeE = oProperty.Value
		end if
	next
	percentUsedE = 100 - (100 * (freeE/sizeE))
next

wqlQuery = "select FreeSpace,Size from Win32_LogicalDisk where Name = 'F:'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "FreeSpace" then
			freeF = oProperty.Value
		elseif oProperty.Name = "Size" then
			sizeF = oProperty.Value
		end if
	next
	percentUsedE = 100 - (100 * (freeF/sizeF))
next


wscript.echo percentUsedC
wscript.echo percentUsedE
wscript.echo percentUsedF

wscript.echo Date() & " " & Time()

wscript.echo "Disk C: Used"
wscript.echo "Disk E: Used"
wscript.echo "Disk F: Used"

