' Run this script at the command prompt by typing
'     cscript processes.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select NumberOfProcesses,NumberOfUsers from Win32_OperatingSystem"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "NumberOfProcesses" then
			numberProcesses = oProperty.Value
		elseif oProperty.Name = "NumberOfUsers" then
			numberUsers = oProperty.Value
		end if
	next
next

wscript.echo numberProcesses
wscript.echo numberUsers

wscript.echo Date() & " " & Time()

wscript.echo "Number of Processes"
wscript.echo "Number of Users"

