strComputer = "apus"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery _
   ("Select * from Win32_Process where name = 'qrx.exe'")
For Each objProcess in colProcesses

    Wscript.Echo "Process: " & objProcess.Name
    sngProcessTime = (CSng(objProcess.KernelModeTime) + _
        CSng(objProcess.UserModeTime)) / 10000000
    Wscript.Echo "Processor Time: " & sngProcessTime
    Wscript.Echo "Process ID: " & objProcess.ProcessID
    Wscript.Echo "Working Set Size: " _
    & objProcess.WorkingSetSize
    Wscript.Echo "Page File Size: " _
    & objProcess.PageFileUsage
    Wscript.Echo "Page Faults: " & objProcess.PageFaults
Next
