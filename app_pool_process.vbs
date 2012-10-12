' app_pool_process.vbs

On Error Resume Next


Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputerName = WshNetwork.ComputerName


Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(WshNetwork.ComputerName)
For Each strComputer In arrComputers
   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process where name = 'w3wp.exe'", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "CreationDate: " & WMIDateStringToDate(objItem.CreationDate)
      WScript.Echo "CSName: " & objItem.CSName
      WScript.Echo "ProcessId: " & objItem.ProcessId
      ' bits to bytes to MB
	  WScript.Echo "PageFileUsage: " & (objItem.PageFileUsage / 1024) / 1024 & " MB"
	  WScript.Echo "WorkingSetSize: " & (objItem.WorkingSetSize / 1024) / 1024 & " MB"
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function