' This script fetches the number of current ASP sessions from 2 different computers
' Syntax: "cscript aspsessions.vbs <machine1> <machine2>"
'

Set objArgs = wscript.Arguments
strComputer1 = objArgs.item(0)
strComputer2 = objArgs.item(1)

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer1 & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_ASP_ActiveServerPages",,48)
For Each objItem in colItems
    intSessions1=objItem.SessionsCurrent
Next

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer2 & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_ASP_ActiveServerPages",,48)
For Each objItem in colItems
    intSessions2=objItem.SessionsCurrent
Next

wscript.echo intSessions1
wscript.echo intSessions2