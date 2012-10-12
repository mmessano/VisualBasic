' Run this script at the command prompt by typing
'     cscript httpUsers.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select AnonymousUsersPersec,NonAnonymousUsersPersec from Win32_PerfRawData_W3SVC_WebService where Name = '_Total'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "AnonymousUsersPersec" then
			usersAnon = oProperty.Value
		elseif oProperty.Name = "NonAnonymousUsersPersec" then
			usersNonAnon = oProperty.Value
		end if
	next
next

wscript.echo usersAnon
wscript.echo usersNonAnon

wscript.echo Date() & " " & Time()

wscript.echo "Number of Anonymous Users"
wscript.echo "Number of Non-Anonymous Users"

