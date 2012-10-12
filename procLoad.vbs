' Run this script at the command prompt by typing
'     cscript procLoad.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select LoadPercentage from Win32_Processor where DeviceID = 'CPU0'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "LoadPercentage" then
			proc0Load = oProperty.Value
		end if
	next
next

wqlQuery = "select LoadPercentage from Win32_Processor where DeviceID = 'CPU1'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "LoadPercentage" then
			proc1Load = oProperty.Value
		end if
	next
next

wscript.echo proc0Load
wscript.echo proc1Load

wscript.echo Date() & " " & Time()

wscript.echo "Proc #1 Load"
wscript.echo "Proc #2 Load"

