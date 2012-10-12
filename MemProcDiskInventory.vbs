 '==========================================================================
'
' NAME: <MemProcDiskInventory.vbs>
'
' AUTHOR: Mark D. MacLachlan , The Spider's Parlor
' URL: http://www.thespidersparlor.com
' DATE  : 2/5/2003
'
' COMMENT: <Inventories computer configurations from a list of computers>
'==========================================================================
on error resume next

set x = getobject(,"excel.application")
r = 2
do until len(x.cells(r, 1).value) = 0
strComputer = x.cells(r, 1).Value

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each OS In colSettings
    x.cells(r, 7).value = OS.Caption
    x.cells(r, 8).value = OS.Version
Next   
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings
x.cells(r, 2).value = objComputer.Name
x.cells(r, 3).value = objComputer.TotalPhysicalMemory /1024\1024+1 & "MB"
Next
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_Processor")
For Each objProcessor in colSettings
x.cells(r, 4).value = objProcessor.Description
Next


Set objWMIService = GetObject("winmgmts:")
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
x.cells(r, 5).value = objLogicalDisk.Size /1024\1024+1 & "MB"
x.cells(r, 6).value = objLogicalDisk.FreeSpace /1024\1024+1 & "MB"

r = r + 1
loop
