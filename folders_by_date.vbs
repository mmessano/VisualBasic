Const LOCAL_TIME = TRUE

Set dtmTargetDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
dtmTargetDate.SetVarDate "3/1/2004", LOCAL_TIME

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory Where " _
        & "CreationDate > '" & dtmTargetDate & "'")

For each objFolder in colFolders
    dtmConvertedDate.Value = objFolder.CreationDate 
    Wscript.Echo objFolder.Name & VbTab & _
        dtmConvertedDate.GetVarDate(LOCAL_TIME)
Next
