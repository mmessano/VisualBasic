'************************************************************************
'*		This script is for listing runnig process' on a remote
'*		computer running Win2K or WinNT w/WMI installed.
'*
'*		Author: Darron Nesbitt
'*		Date: 	8/14/2001
'*
'************************************************************************
Function PadStr(Str,Pad)
	Dim CLen

	CLen = Pad - Len(Str)

	PadStr = Str & Space(CLen)
End Function

Dim strComputerName ' The Computer Name to be queried via WMI
Dim strWinMgt		' The WMI management String
Dim Processes		'Hold Processes

On Error Resume Next

'get computer's name or ip address
strComputerName = ucase(InputBox("Enter the remote computers name or IP","Computer Name/IP"))

strWinMgt = "winmgmts://" & strComputerName

'
' Get Computer/User Info
'
 Set CompSysSet = GetObject(strWinMgt).ExecQuery("select * from Win32_ComputerSystem")
 for each CompSys in CompSysSet
         strDescription = CompSys.Description
         strModel       = CompSys.Model
         strName        = CompSys.Name
	 strManufacturer	= CompSys.Manufacturer
	 strUserName		= CompSys.UserName
 next

 CompInfo = "Computer Information" & VBCrLf & VBCrLf
 CompInfo = CompInfo & "Computer Name: " & strName & VBTab & "User: " & strUserName & VBCrLf

'connect to processes
Set Processes = GetObject(strWinMgt).ExecQuery ("Select * from Win32_Process")

'Setup columns for Process Name,PID, & Owner
ProcessInfo = "Process Name" & "  ,  " & "Process ID" & VBCRLF & VBCRLF

'Loop through process
	for each Process in Processes

		PCap = Process.Caption
		PID = Process.ProcessID

		ProcessInfo = ProcessInfo & PCap & "  ,  " & PID & VBCRLF
	next

	'display info about computer and process
	RetVal = MsgBox (CompInfo & VBCRLF & VBCRLF & _
			 ProcessInfo,VBOKOnly,strName & " - List Proccess'")