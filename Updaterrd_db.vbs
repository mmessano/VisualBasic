'Objective: Update an RRD database with current CPULoad information of the server
'Created by: MAK
'Date Apr 23, 2005
'Usage: cscript //b //nologo Updaterrd_db.vbs atdbqa

Set WshShell = WScript.CreateObject("WScript.Shell")

Set objArgs = WScript.Arguments
Computer=objArgs(0)
Set procset = 
	GetObject("winmgmts:{impersonationLevel=impersonate}!\\" 
	& Computer & "\root\cimv2").InstancesOf 
	("Win32_Processor")


for each System in ProcSet
query ="rrdtool update "&computer& "_" & system.deviceid & ".rrd " 
	& UDate(getutc(now())) &":" &system.LoadPercentage 
Return = WshShell.Run(Query, 1)
next

function UDate(oldDate)
  UDate = DateDiff("s", "01/01/1970 00:00:00", oldDate)
end function

function getutc(mydate)

    od = mydate
    set oShell = CreateObject("WScript.Shell") 
    atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias" 
    offsetMin = oShell.RegRead(atb) 
    nd = dateadd("n", offsetMin, od) 
    wscript.echo nd    
    'Response.Write("Current = " & od & "<br>UTC = " & nd) 
    getutc= nd
end function