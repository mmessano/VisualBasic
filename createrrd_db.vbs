'Objective: TO create one RRD database for every processor for a given server name
'Created by:MAK
'Date: Apr 23, 2005
'Usage: cscript //b //nologo createrrd_db.vbs atdbqa

Set WshShell = WScript.CreateObject("WScript.Shell")

Set objArgs = WScript.Arguments
Computer=objArgs(0)

Set procset = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Computer & "\root\cimv2").InstancesOf ("Win32_Processor")

for each System in ProcSet
query ="rrdtool create " & Computer &"_"& system.deviceid &".rrd --start " & UDate(getutc(now())) & " --step 300 DS:LOAD1:GAUGE:600:-1:100 RRA:AVERAGE:0.5:1:1200"
wscript.echo query

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
'    wscript.echo nd    
    'Response.Write("Current = " & od & "<br>UTC = " & nd) 
    getutc= nd
end function
