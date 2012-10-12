' Run this script at the command prompt by typing
'     cscript usedMem.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select FreeVirtualMemory,FreePhysicalMemory from Win32_OperatingSystem"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "FreeVirtualMemory" then
			freeVirtual = oProperty.Value
		elseif oProperty.Name = "FreePhysicalMemory" then
			freePhysical = oProperty.Value
		end if
	next
next

wqlQuery = "select TotalVirtualMemory,TotalPhysicalMemory from Win32_LogicalMemoryConfiguration"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "TotalVirtualMemory" then
			totalVirtual = oProperty.Value
		elseif oProperty.Name = "TotalPhysicalMemory" then
			totalPhysical = oProperty.Value
		end if
	next
next

percentVirtualUsed = 100 - (100 * (freeVirtual / totalVirtual))
percentPhysicalUsed = 100 - (100 * (freePhysical / totalPhysical))

wscript.echo percentVirtualUsed
wscript.echo percentPhysicalUsed

wscript.echo Date() & " " & Time()

wscript.echo "Percent Used Virutal Memory"
wscript.echo "Percent Used Physical Memory"

