Const LOCAL_TIME = TRUE

Set dtmTargetDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
dtmTargetDate.SetVarDate "8/1/2004", LOCAL_TIME

strComputer = "paweb5"
'strPath = "\relateprod\PrimeAlliance"
Wscript.Echo strPath
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory Where " & "CreationDate < '" & dtmTargetDate & "' AND Drive = E")

For each objFolder in colFolders
    dtmConvertedDate.Value = objFolder.CreationDate
    Wscript.Echo objFolder.Name & VbTab & _
        dtmConvertedDate.GetVarDate(LOCAL_TIME)
Next
