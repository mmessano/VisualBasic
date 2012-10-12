'
' Copyright (c) Microsoft Corporation.  All rights reserved.
'
' VBScript Source File 
'
' Script Name: IIsApp.vbs
'

Option Explicit
On Error Resume Next

' Error codes
Const ERR_OK              = 0
Const ERR_GENERAL_FAILURE = 1

'''''''''''''''''''''
' Messages
Const L_Gen_ErrorMessage               = "%1 : %2"
Const L_CmdLib_ErrorMessage            = "Could not create an instance of the CmdLib object."
Const L_ChkCmdLibReg_ErrorMessage      = "Please register the Microsoft.CmdLib component."
Const L_ScriptHelper_ErrorMessage      = "Could not create an instance of the IIsScriptHelper object."
Const L_ChkScpHelperReg_ErrorMessage   = "Please, check if the Microsoft.IIsScriptHelper is registered."
Const L_InvalidSwitch_ErrorMessage     = "Invalid switch: %1"
Const L_NotEnoughParams_ErrorMessage   = "Not enough parameters."
Const L_Query_ErrorMessage             = "Error executing query"
Const L_Serving_Message                = "The following W3WP.exe processes are serving AppPool: %1"
Const L_NoW3_ErrorMessage              = "Error - no w3wp.exe processes are running at this time"
Const L_PID_Message                    = "W3WP.exe PID: %1"
Const L_NoResults_ErrorMessage         = "Error - no results"
Const L_APID_Message                   = "W3WP.exe PID: %1   AppPoolId: %2"
Const L_NotW3_ErrorMessage             = "ERROR: ProcessId specified is NOT an instance of W3WP.exe - EXITING"
Const L_PIDNotValid_ErrorMessage       = "ERROR: PID is not valid"
Const L_Recycled_Message               = "Application pool '%1' recycled successfully."
Const L_PoolDoesntExist_ErrorMessage   = "Application pool '%1' does not exist."
Const L_Recycle_ErrorMessage           = "Failed to recycle application pool '%1'."
Const L_WMIConnect_ErrorMessage        = "Could not connect to WMI provider."
Const L_PIDNotFound_ErrorMessage       = "Process ID %1 not found."

'''''''''''''''''''''
' Help
Const L_Empty_Text     = ""

' General help messages
Const L_SeeHelp_Message         = "Type IIsApp /? for help."
 
Const L_Help_HELP_General01_Text = "Description: list IIS application pools and associated worker processes."
Const L_Help_HELP_General02_Text = "             Recycle application pools."
Const L_Help_HELP_General03_Text = "Syntax: IIsApp.vbs [{ /a <app_pool_id> | /p <pid> } [/r] ]"
Const L_Help_HELP_General04_Text = "Parameters:"
Const L_Help_HELP_General05_Text = ""
Const L_Help_HELP_General06_Text = "Value              Description"
Const L_Help_HELP_General07_Text = "/a <app_pool_id>   Specify an application pool by name. Surround"
Const L_Help_HELP_General08_Text = "                   <app_pool_id> with quotes if it contains spaces."
Const L_Help_HELP_General09_Text = "                   If used alone without an accompanying action,"
Const L_Help_HELP_General10_Text = "                   IIsApp.vbs will report PIDs of currently running"
Const L_Help_HELP_General11_Text = "                   w3wp.exe processes serving pool <app_pool_id>."
Const L_Help_HELP_General12_Text = "/p <pid>           Specify a process by process ID. If used alone"
Const L_Help_HELP_General13_Text = "                   without an accompanying action, IIsApp.vbs will"
Const L_Help_HELP_General14_Text = "                   report the AppPoolId of the w3wp process specified"
Const L_Help_HELP_General15_Text = "                   by <pid>. When a PID is specified with /r, that PID"
Const L_Help_HELP_General16_Text = "                   is mapped to an application pool and the action is"
Const L_Help_HELP_General17_Text = "                   taken upon the application pool. If a PID is given"
Const L_Help_HELP_General18_Text = "                   for a web garden, i.e. an application pool served"
Const L_Help_HELP_General19_Text = "                   by more than one w3wp, then all w3wp’s for that"
Const L_Help_HELP_General20_Text = "                   application pool will be acted upon."
Const L_Help_HELP_General21_Text = "/r                 Recycles the application pool."
Const L_Help_HELP_General22_Text = "DEFAULT: no switches will print out the PID and AppPoolId."
Const L_Help_HELP_General23_Text = "Examples:"
Const L_Help_HELP_General24_Text = "IIsApp"
Const L_Help_HELP_General25_Text = "IIsApp /p 2368"
Const L_Help_HELP_General26_Text = "IIsApp /a DefaultAppPool /r"
Const L_Help_HELP_General27_Text = "IIsApp /p 2368 /r"

''''''''''''''''''''''''
' Operation codes
Const OPER_BY_NAME = 1
Const OPER_BY_PID  = 2
Const OPER_ALL     = 3

'
' Main block
'
Dim oScriptHelper, oCmdLib
Dim intOperation, intResult
Dim strCmdLineOptions
Dim oError
Dim aArgs
Dim apoolID, PID
Dim oProviderObj
Dim bRecycle

' Default
intOperation = OPER_ALL
bRecycle = False

Const wmiConnect  = "winmgmts:{(debug)}:/root/cimv2"
Const queryString = "select * from Win32_Process where Name='w3wp.exe'"
Const pidQuery    = "select * from Win32_Process where ProcessId="

' get NT WMI provider
Set oProviderObj = GetObject(wmiConnect)

' Instantiate the CmdLib for output string formatting
Set oCmdLib = CreateObject("Microsoft.CmdLib")
If Err.Number <> 0 Then
    WScript.Echo L_CmdLib_ErrorMessage
    WScript.Echo L_ChkCmdLibReg_ErrorMessage    
    WScript.Quit(ERR_GENERAL_FAILURE)
End If
Set oCmdLib.ScriptingHost = WScript.Application

' Instantiate script helper object
Set oScriptHelper = CreateObject("Microsoft.IIsScriptHelper")
If Err.Number <> 0 Then
    WScript.Echo L_ScriptHelper_ErrorMessage
    WScript.Echo L_ChkScpHelperReg_ErrorMessage    
    WScript.Quit(ERR_GENERAL_FAILURE)
End If

Set oScriptHelper.ScriptHost = WScript

' Check if we are being run with cscript.exe instead of wscript.exe
oScriptHelper.CheckScriptEngine

' Command Line parsing
Dim argObj, arg
Set argObj = WScript.Arguments

strCmdLineOptions = "[a:a:1;r:recycle:0];[p:p:1;r:recycle:0]"

If argObj.Named.Count > 0 Then
    Set oError = oScriptHelper.ParseCmdLineOptions(strCmdLineOptions)

    If Not oError Is Nothing Then
        If oError.ErrorCode = oScriptHelper.ERROR_NOT_ENOUGH_ARGS Then
            ' Not enough arguments for a specified switch
            WScript.Echo L_NotEnoughParams_ErrorMessage
            WScript.Echo L_SeeHelp_Message
        Else
            ' Invalid switch
            oCmdLib.vbPrintf L_InvalidSwitch_ErrorMessage, Array(oError.SwitchName)
      	    WScript.Echo L_SeeHelp_Message
        End If
        
            WScript.Quit(ERR_GENERAL_FAILURE)
    End If

    If oScriptHelper.GlobalHelpRequested Then
        DisplayHelpMessage
        WScript.Quit(ERR_OK)
    End If

    For Each arg In oScriptHelper.Switches
        Select Case arg
            Case "a"
                apoolID = oScriptHelper.GetSwitch(arg)
                intOperation = OPER_BY_NAME
                
            Case "p"
                PID = oScriptHelper.GetSwitch(arg)
                intOperation = OPER_BY_PID
                
            Case "r"
                bRecycle = True
        End Select
    Next

End If

' Choose operation
Select Case intOperation
	Case OPER_BY_NAME
		intResult = GetByPool(apoolID)
		
	Case OPER_BY_PID
		intResult = GetByPid(PID)

    Case OPER_ALL
        intResult = GetAllW3WP()
End Select

' Return value to command processor
WScript.Quit(intResult)

'''''''''''''''''''''''''
' End Of Main Block
'''''''''''''''''''''

'''''''''''''''''''''''''''
' DisplayHelpMessage
'''''''''''''''''''''''''''
Sub DisplayHelpMessage()
    WScript.Echo L_Help_HELP_General01_Text
    WScript.Echo L_Help_HELP_General02_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General03_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General04_Text
    WScript.Echo L_Help_HELP_General05_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Help_HELP_General11_Text
    WScript.Echo L_Help_HELP_General12_Text
    WScript.Echo L_Help_HELP_General13_Text
    WScript.Echo L_Help_HELP_General14_Text
    WScript.Echo L_Help_HELP_General15_Text
    WScript.Echo L_Help_HELP_General16_Text
    WScript.Echo L_Help_HELP_General17_Text
    WScript.Echo L_Help_HELP_General18_Text
    WScript.Echo L_Help_HELP_General19_Text
    WScript.Echo L_Help_HELP_General20_Text
    WScript.Echo L_Help_HELP_General21_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General22_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General23_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General24_Text
    WScript.Echo L_Help_HELP_General25_Text
    WScript.Echo L_Help_HELP_General26_Text
    WScript.Echo L_Help_HELP_General27_Text
End Sub

Function GetAppPoolId(strArg)
	Dim Submatches
	Dim strPoolId
	Dim re
    Dim Matches

    On Error Resume Next

	Set re = New RegExp
	re.Pattern = "-ap ""(.+)"""
	re.IgnoreCase = True
	Set Matches = re.Execute(strArg)
	Set SubMatches = Matches(0).Submatches
	strPoolId = Submatches(0)
	
	GetAppPoolId = strPoolId
End Function

Function GetByPool(strPoolName)
	Dim W3WPList
	Dim strQuery
    Dim W3WP

    On Error Resume Next

    If bRecycle Then
		GetByPool = RecycleAppPool(strPoolName)
		Exit Function
	End If

	strQuery = queryString
	Set W3WPList = oProviderObj.ExecQuery(strQuery)
	If (Err.Number <> 0) Then
		WScript.Echo L_Query_ErrorMessage
        oCmdLib.vbPrintf L_Gen_ErrorMessage, Array(Hex(Err.Number), Err.Description)
		GetByPid = 2
	Else
        oCmdLib.vbPrintf L_Serving_Message, Array(strPoolName)
		If (W3WPList.Count < 1) Then
			WScript.Echo L_NoW3_ErrorMessage
			GetByPool = 1
		Else
			For Each W3WP In W3WPList
				If (UCase(GetAppPoolId(W3WP.CommandLine)) = UCase(strPoolName)) Then
                    oCmdLib.vbPrintf L_PID_Message, Array(W3WP.ProcessId)
				End If
			Next
			GetByPool = 0
		End If
	End If
End Function

Function GetByPid(pid)
    Dim result, poolName

    On Error Resume Next

    result = GetPoolByPid(pid, poolName)
    Select Case result
        ' Successful case
        Case 0
            If bRecycle Then
				result = RecycleAppPool(poolName)
			Else
	            oCmdLib.vbPrintf L_APID_Message, Array(pid, poolName)
            End If
        
        ' No process with such PID was found
        Case &H80070002
            oCmdLib.vbPrintf L_PIDNotFound_ErrorMessage, Array(pid)

        ' Pid should be a number
        Case &H80070057
    		WScript.Echo L_PIDNotValid_ErrorMessage
        
        ' Process found, but it is not a worker process
        Case &H8007000D
    		WScript.Echo L_NotW3_ErrorMessage
    		
    	' Some error occurred
    	Case Default
			WScript.Echo L_Query_ErrorMessage
            oCmdLib.vbPrintf L_Gen_ErrorMessage, Array(Hex(Err.Number), Err.Description)
    End Select

    GetByPid = result
End Function
    
Function GetPoolByPid( pid, ByRef poolName )
	Dim W3WPList
    Dim strQuery
    Dim W3WP

    On Error Resume Next

    poolName = ""
	If IsNumeric(pid) Then
		strQuery = pidQuery & pid
	
		Set W3WPList = oProviderObj.ExecQuery(strQuery)
		If Err.Number <> 0 Then
			GetPoolByPid = Err.Number
		Else
			If W3WPList.Count < 1 Then
				GetPoolByPid = &H80070002
			Else
				For Each W3WP In W3WPList
					If UCase(W3WP.Name) = "W3WP.EXE" Then
						poolName = GetAppPoolId(W3WP.CommandLine)
						GetPoolByPid = 0
					Else
						GetPoolByPid = &H8007000D
					End If
				Next
			End If
		End If
	Else
		GetPoolByPid = &H80070057
	End If
End Function

Function GetAllW3WP()
	Dim W3WPList
	Dim strQuery
    Dim W3WP

    On Error Resume Next

	strQuery = queryString
	Set W3WPList = oProviderObj.ExecQuery(strQuery)
	If (Err.Number <> 0) Then
		WScript.Echo L_Query_ErrorMessage
        oCmdLib.vbPrintf L_Gen_ErrorMessage, Array(Hex(Err.Number), Err.Description)
		GetByPid = 2
	Else
		If (W3WPList.Count < 1) Then
			WScript.Echo L_NoResults_ErrorMessage
			GetAllW3WP = 2
		Else
			For Each W3WP In W3WPList
                oCmdLib.vbPrintf L_APID_Message, Array(W3WP.ProcessId, GetAppPoolId(W3WP.CommandLine))
			Next
			GetAllW3WP = 0
		End If
	End If
End Function

Function RecycleAppPool(apoolID)
    Dim PoolObj
    
    On Error Resume Next
    
    ' Initializes authentication with remote machine
    intResult = oScriptHelper.InitAuthentication(".", "", "")
    If intResult <> 0 Then
        RecycleAppPool = intResult
        Exit Function
    End If

    oScriptHelper.WMIConnect
    If Err.Number Then
        WScript.Echo L_WMIConnect_ErrorMessage
        oCmdLib.vbPrintf L_Gen_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        RecycleAppPool = Err.Number
        Exit Function
    End If

    Set PoolObj = oScriptHelper.ProviderObj.Get("IIsApplicationPool='W3SVC/AppPools/" & apoolID & "'")
    If Err.Number Then
        If Err.Number = &H80070003 Then
            oCmdLib.vbPrintf L_PoolDoesntExist_ErrorMessage, Array(apoolID)
        Else
            oCmdLib.vbPrintf L_Gen_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        End If
        
        RecycleAppPool = Err.Number
        Exit Function
    End If
    
    PoolObj.Recycle
    If Err.Number Then
        oCmdLib.vbPrintf L_Recycle_ErrorMessage, Array(apoolID)
        oCmdLib.vbPrintf L_Gen_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        RecycleAppPool = Err.Number
        Exit Function
    End If
    
    oCmdLib.vbPrintf L_Recycled_Message, Array(apoolID)
    RecycleAppPool = 0
End Function
