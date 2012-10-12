On Error Resume Next
strComputer = "paweb1"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DCOMApplicationSetting",,48)
For Each objItem in colItems
    Wscript.Echo "AppID: " & objItem.AppID
    Wscript.Echo "AuthenticationLevel: " & objItem.AuthenticationLevel
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CustomSurrogate: " & objItem.CustomSurrogate
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "EnableAtStorageActivation: " & objItem.EnableAtStorageActivation
    Wscript.Echo "LocalService: " & objItem.LocalService
    Wscript.Echo "RemoteServerName: " & objItem.RemoteServerName
    Wscript.Echo "RunAsUser: " & objItem.RunAsUser
    Wscript.Echo "ServiceParameters: " & objItem.ServiceParameters
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "UseSurrogate: " & objItem.UseSurrogate
    Wscript.Echo CrLf
Next
