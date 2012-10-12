'
'
On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("apus")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process where Name='qrx.exe'", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "##### Begin process #####"
	  WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "CommandLine: " & objItem.CommandLine
'      WScript.Echo "CreationClassName: " & objItem.CreationClassName
      WScript.Echo "CreationDate: " & WMIDateStringToDate(objItem.CreationDate)
'      WScript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
      WScript.Echo "CSName: " & objItem.CSName
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "ExecutablePath: " & objItem.ExecutablePath
'      WScript.Echo "ExecutionState: " & objItem.ExecutionState
'      WScript.Echo "Handle: " & objItem.Handle
'      WScript.Echo "HandleCount: " & objItem.HandleCount
'      WScript.Echo "InstallDate: " & WMIDateStringToDate(objItem.InstallDate)
'      WScript.Echo "KernelModeTime: " & objItem.KernelModeTime
      WScript.Echo "MaximumWorkingSetSize: " & objItem.MaximumWorkingSetSize
      WScript.Echo "MinimumWorkingSetSize: " & objItem.MinimumWorkingSetSize
      WScript.Echo "Name: " & objItem.Name
'      WScript.Echo "OSCreationClassName: " & objItem.OSCreationClassName
'      WScript.Echo "OSName: " & objItem.OSName
'      WScript.Echo "OtherOperationCount: " & objItem.OtherOperationCount
'      WScript.Echo "OtherTransferCount: " & objItem.OtherTransferCount
      WScript.Echo "PageFaults: " & objItem.PageFaults
      WScript.Echo "PageFileUsage: " & objItem.PageFileUsage
      WScript.Echo "ParentProcessId: " & objItem.ParentProcessId
      WScript.Echo "PeakPageFileUsage: " & objItem.PeakPageFileUsage
      WScript.Echo "PeakVirtualSize: " & objItem.PeakVirtualSize
      WScript.Echo "PeakWorkingSetSize: " & objItem.PeakWorkingSetSize
'      WScript.Echo "Priority: " & objItem.Priority
      WScript.Echo "PrivatePageCount: " & objItem.PrivatePageCount
      WScript.Echo "ProcessId: " & objItem.ProcessId
'      WScript.Echo "QuotaNonPagedPoolUsage: " & objItem.QuotaNonPagedPoolUsage
'      WScript.Echo "QuotaPagedPoolUsage: " & objItem.QuotaPagedPoolUsage
'      WScript.Echo "QuotaPeakNonPagedPoolUsage: " & objItem.QuotaPeakNonPagedPoolUsage
'      WScript.Echo "QuotaPeakPagedPoolUsage: " & objItem.QuotaPeakPagedPoolUsage
'      WScript.Echo "ReadOperationCount: " & objItem.ReadOperationCount
'      WScript.Echo "ReadTransferCount: " & objItem.ReadTransferCount
      WScript.Echo "SessionId: " & objItem.SessionId
'      WScript.Echo "Status: " & objItem.Status
'      WScript.Echo "TerminationDate: " & WMIDateStringToDate(objItem.TerminationDate)
'      WScript.Echo "ThreadCount: " & objItem.ThreadCount
'      WScript.Echo "UserModeTime: " & objItem.UserModeTime
      WScript.Echo "VirtualSize: " & objItem.VirtualSize
'      WScript.Echo "WindowsVersion: " & objItem.WindowsVersion
      WScript.Echo "WorkingSetSize: " & objItem.WorkingSetSize
'      WScript.Echo "WriteOperationCount: " & objItem.WriteOperationCount
'      WScript.Echo "WriteTransferCount: " & objItem.WriteTransferCount
      WScript.Echo "##### End process #####"
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function



