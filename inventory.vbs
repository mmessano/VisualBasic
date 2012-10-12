' Usage: cscript inventory.vbs ComputerName PathToSaveFile
'Copyright 2002 MouseTrax Computing Solutions
'http://www.mousetrax.com
'
'Author: Greg Chapman
'----------------------------------------------------------------------

'**********************************************************************
'Constants for log file operations                                   '*
'**********************************************************************
CONST ForReading = 1, ForWriting = 2, ForAppending = 8               '*
'**********************************************************************

Dim objWbem, objWbemObjectSet, objWbemObject, objWshNetwork, objOS, fso
Dim strClassName, strPropertyName, strComputerName, CPUType, num, Gb 
Dim Iyear, InMonth, IMonth, IDate, months
Dim strTotalSystem, cstrQ, cstrQCQ, cstrSeparator

cstrQ = Chr(34)
cstrQCQ = Chr(34) & "," & Chr(34)
cstrSeparator = String(72,"=")

' Check command-line argument collection. If the user supplied an 
' argument, copy to strComputerName. If the argument contains a 
' question mark, call the Usage subroutine and exit. If the user 
' didn't supply an argument, set strComputerName to the local
' machine name.

If(WScript.Arguments.Count) = 2 Then
  strComputerName = WScript.Arguments.Item(0)
  If(InStr(strComputerName, "?")) Then
    Call Usage()
  End If

  strSaveDir= Wscript.Arguments.Item(1)
  If Right(strSaveDir,1) <> "\" Then
    strSaveDir=strSaveDir & "\"
  End If
ElseIf Wscript.Arguments.Count = 1 Then
   strComputerName = wscript.arguments.item(0)
   wscript.echo strComputerName
   strSaveDir=ExecutingFrom 
Else
   Set objWshNetwork = WScript.CreateObject("WScript.Network")
   strComputerName = objWshNetwork.ComputerName
   wscript.echo strComputerName
   Set objWshNetwork = Nothing
   strSaveDir=ExecutingFrom  
End If

Set fso=CreateObject("Scripting.FileSystemObject")
LogFile=strSaveDir & strComputerName & ".txt"

ReachResult=reachable(strComputerName)

If (Instr(1,LCase(ReachResult),"unknown host") > 0) OR  _
  (Instr(1,LCase(ReachResult),"could not find host")>0) Then
  'write to unknown_hosts.txt
  WriteInv strSaveDir & "Unknown_Hosts.txt", strComputerName
  wscript.quit
ElseIf (Instr(1,ReachResult,"unreachable") > 0) OR _
  (Instr(1,ReachResult,"not reachable") > 0) OR _
  (ReachResult="") Then
  'write to Unreachable_Hosts.txt
  WriteInv strSaveDir & "Unreachable_Hosts.txt", strComputerName
  wscript.quit
End If

strTotalSystem=""
if ReachResult <> "" then
  strTotalSystem="[Network Information]" & vbNewLine
 	strTotalSystem = strTotalSystem & "IP Address = " & ReachResult

  strNicResults=GetNicsFinal(strComputerName)
  If strNicResults = "Investigate" Then
    wscript.quit
  Else
	  strTotalSystem = strTotalSystem & vbNewLine & strNicResults & vbNewLine 
  End If
  strTotalSystem = strTotalSystem & vbNewLine & "[System]"
	strTotalSystem = strTotalSystem & vbNewLine & strComputerName 
  strTotalSystem = strTotalSystem & vbnewLine & "SID = " & _
  WMIGetSid(strComputerName)
	strTotalSystem = strTotalSystem & vbNewLine & _
  getInventory (strComputerName) & vbNewLine
	strTotalSystem = strTotalSystem & vbNewLine & _
  "[Installed Software]"
	strTotalSystem = strTotalSystem & vbNewLine & _
  GetInstalledAppsReg (strComputerName) & vbNewLine
	strTotalSystem = strTotalSystem & vbNewLine & _
  "[Running Tasks]" & vbNewLine & _
  "Process Name" & vbTab & "Process ID" & vbTab & "Phys Mem" & vbTab & "Vir Mem"
	strTotalSystem = strTotalSystem & vbNewLine & _
  ProcessProcesses (strComputerName) & vbNewLine
	strTotalSystem = strTotalSystem & vbNewLine & _
  "[Running Services]"
	strTotalSystem = strTotalSystem & vbNewLine & _
  ProcessServices (strComputerName)
else
	strTotalSystem = strTotalSystem & vbNewLine & _
  strComputerName & " not reachable"
end if

LogAction strTotalSystem
'====================================================================
Function getInventory (StrComputer)

On Error Resume Next

	Set objWbem = _
  GetObject("winmgmts:{ImpersonationLevel = Impersonate}//" _
  & StrComputer)

'  Get the Vendor, Name, ID  
	strClassName = "Win32_ComputerSystemProduct"
	Set objWbemObjectSet = objWbem.ExecQuery _
  ("Select Vendor, Name, IdentifyingNumber From " & _ 
		strClassName)
	strSystemProduct = ""
	For Each objWbemObject In objWbemObjectSet
   		strVendor= "PC Vendor= " & _
      objWbemObject.Properties_("Vendor") & _
      vbNewLine
		If Err <> 0 Then
			strVendor="PC Vendor = N/A" & vbNewLine
			Err.Clear
			Err.Number = 0
			WriteInv strSaveDir & "UpdateWMI.txt", strComputerName
			Exit Function
      			'wscript.quit
		End If

   	strVenName= "Vendor Model= " & objWbemObject.Properties_("Name")& _
    vbNewLine
		If Err <> 0 Then
			strVenName= "Name = Not Available" & vbNewLine
			Err.Clear
			Err.Number = 0
		End If

   		strSN= "Vendor ID Serial Number= " & _
      objWbemObject.Properties_("IdentifyingNumber")
		If Err <> 0 Then
			strSN= "ID Serial Number = Not Available"
			Err.Clear
			Err.Number = 0
		End If
    If GetInventory <> "" Then
  		GetInventory=GetInventory & vbNewLine & strSystemProduct & _
      strVendor & strVenName & strSN 
    Else
      GetInventory=strSystemProduct & strVendor & strVenName & _
      strSN 
    End If
	Next

'  Get the Family, CurrentClockSpeed, Description  
	strClassName = "Win32_Processor"
	strPropertyName = "Family"
	Set objWbemObjectSet = _
  objWbem.ExecQuery _
  ("Select Family, CurrentClockSpeed, Description From " & _ 
		strClassName)
CPUType = Array("Other","Unknown","8086","80286","80386", _
"80486","8087","80287","80387","80487","Pentium Family",_
 "Pentium Pro","Pentium II","Pentium MMX","Celeron", _
 "Pentium II Xeon","Pentium III","M1 Family","M2 Family",_
 "K5 Family","K6 Family","K6-2","K6-III","Athlon","Power PC Family", _
 "Power PC 601","Power PC 603","Power PC 603+",_
 "Athlon 1800 XP","Alpha Family","MIPS Family","SPARC Family", _
 "68040","68xxx Family","68000","68010","68020","68030",_
 "Hobbit Family","Weitek","PA-RISC Family","V30 Family", _
 "Pentium III Xeon","AS400 Family","IBM390 Family","i860",_
 "i960","SH-3","SH-4","ARM","StrongARM","6x86","MediaGX", _
 "MII","WinChip")
	strProcProfile=""
	For Each objWbemObject In objWbemObjectSet
  		strProcFamily= "Proc Family = " & _
      CPUType(objWbemObject.Properties_("Family")-1)
		If Err <> 0 Then
			strProcFamily= "Proc Family = Not Available"
			Err.Clear
			Err.Number = 0
		End If
		
  		strClock= "Clock Speed = " & _
      objWbemObject.Properties_("CurrentClockSpeed")
		If Err <> 0 Then
			strClock= "Clock Speed = Not Available"
			Err.Clear
			Err.Number = 0
		End If

  		strDescription= "Description = " & _
      objWbemObject.Properties_("Description")
		If Err <> 0 Then
			strDescription= "Description = Not Available"
			Err.Clear
			Err.Number = 0
		End If
		strProcProfile=strProcProfile & strProcFamily & " " & _
    strClock & " " & strDescription & vbNewLine
	Next
	GetInventory = GetInventory & vbNewLine & strProcProfile

'  Get the Memory  
	strClassName = "Win32_LogicalMemoryConfiguration"
	strPropertyName = "TotalPhysicalMemory"
	Set objWbemObjectSet = objWbem.ExecQuery("Select " & _
  strPropertyName & " From " & _ 
		strClassName)
	For Each objWbemObject In objWbemObjectSet
   		strMemory = "Memory = " & _
      FormatNumber(objWbemObject.Properties_(strPropertyName)/1024,0) _
      & "(Mb)"
		If Err <> 0 Then
			strMemory= "Memory = Not Available"
			Err.Clear
			Err.Number = 0
		End If
	Next
  GetInventory = GetInventory & vbNewLine & strMemory

'  Get the Release Date  
	strClassName = "Win32_BIOS"
	strPropertyName = "ReleaseDate"
	Set objWbemObjectSet = objWbem.ExecQuery("Select " & _
  strPropertyName & " From " & _ 
		strClassName)
	For Each objWbemObject In objWbemObjectSet
		Iyear = left(objWbemObject.Properties_(strPropertyName),4)
		InMonth = mid(objWbemObject.Properties_(strPropertyName),5,2)
		IDate = mid(objWbemObject.Properties_(strPropertyName),7,2)
		months = array("blank","Jan","Feb","Mar","Apr","May","Jun", _
    "Jul","Aug","Sep","Oct","Nov","Dec")
		strBios = "Date BIOS Released = " & IDate & " " & _
    Months(InMonth) & " " & Iyear
	Next
  GetInventory = GetInventory & vbNewLine & strBIOS


'  Get the Install Date  
	strClassName = "Win32_OperatingSystem"
	strPropertyName = "InstallDate"
	Set objWbemObjectSet = objWbem.ExecQuery("Select " & _
  strPropertyName & " From " & _ 
		strClassName)
	For Each objWbemObject In objWbemObjectSet
		Iyear = left(objWbemObject.Properties_(strPropertyName),4)
		InMonth = mid(objWbemObject.Properties_(strPropertyName),5,2)
		IDate = mid(objWbemObject.Properties_(strPropertyName),7,2)
		strOSInst = "Date OS installed = " & IDate & " " & _
    Months(InMonth) & " " & Iyear
		GetInventory = GetInventory & vbNewLine & strOSInst
	Next
For Each objOS in GetObject("winmgmts:\\" & strComputer). _
InstancesOf ("Win32_OperatingSystem")
	strOSName= "OS Name = " & objOS.Caption & vbNewLine & _
  "OS Version = " & objOS.Version 
	If Err <> 0 Then
        strOSName= "Name = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strOSName

	strSysDesc= "System Description = " & objOS.Description 
    If Err <> 0 Then
        strSysDesc= "System Description = Not Available" 
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine & strSysDesc

	strRegUser= "Registered User = " & objOS.RegisteredUser  
    If Err <> 0 Then
        strRegUser="Registered User = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine & strRegUser

	strOrg= "Organization = " & objOS.Organization 
    If Err <> 0 Then
        strOrg= "Organization = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strOrg

	strLicUser= "Number of Licensed Users = " & _
  objOs.NUmberOfLicensedUsers 
    If Err <> 0 Then
        strLicUser= "Number of Licensed Users = Not Available" 
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strLicUser

	strNumUsers= "Number of Users = " & objOS.NumberOfUsers 
    If Err <> 0 Then
        strNumUsers= "Numer of Users = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strNumUsers

	strOSSN= "Serial Number = " & objOS.SerialNumber 
    If Err <> 0 Then
        strOSSN= "Serial Number = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strOSSN

	strBuild= "Build Type = " & objOS.BuildType 
    If Err <> 0 Then
        strBuild= "Build Type = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strBuild

	strSP= "OS Service Pack = " & objOS.ServicePackMajorVersion & _
  "." & objOS.ServicePackMinorVersion 
    If Err <> 0 Then
        strSP= "Service Pack = Not Available" 
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strSP

	strInstType= "Product Installation Type = " & objOS.OSProductSuite 
    If Err <> 0 Then
        strInstType= "Product Installation Type = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strInstType

	strEncr= "System Encryption Level = " & objOS.EncryptionLevel 
    If Err <> 0 Then
        strEncr= "System Encryption Level = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strEncr

	strPhysMem= "Physical Memory = " & _
  FormatNumber((objOS.TotalVisibleMemorySize /1024),2) & " MBytes"
    If Err <> 0 Then
        strPhysMem= "Physical Memory = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strPhysMem

	strFreeMem= "Free Physical Memory = " & _
  FormatNumber((objOS.FreePhysicalMemory /1024),2) & " MBytes"
    If Err <> 0 Then
        strFreeMem= "Free Physical Memory = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strFreeMem

	strStatus= "System Status = " & objOS.Status 
    If Err <> 0 Then
        strStatus= "System Status = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strStatus

	strBoost= "Foreground Application Boost = " & _
  objOS.ForegroundApplicationBoost 
    If Err <> 0 Then
        strBoost= "Foreground Application Boost = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strBoost

	strBootDrv= "Boot Device = " & objOS.BootDevice 
    If Err <> 0 Then
        strBootDrv= "Boot Device = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strBootDrv

	strSysDir= "System Directory = " & objOS.SystemDirectory 
    If Err <> 0 Then
        strSysDir= "System Directory = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine &  strSysDir

	strLastBoot= "Last Boot Time = " & _
  ConvWbemTime(objOS.LastBootUpTime)
    If Err <> 0 Then
        strLastBoot= "Last Boot Time = Not Available"
        Err.Clear
        Err.Number = 0
    End If
	GetInventory = GetInventory & vbNewLine & strLastBoot

Next


'  Get the DiskDrive  
	strClassName = "Win32_DiskDrive"
	strPropertyName = "Size"
	Set objWbemObjectSet = objWbem.ExecQuery("Select " & _
  strPropertyName & " From " & _ 
		strClassName)
	For Each objWbemObject In objWbemObjectSet
		Gb = _
    FormatNumber(objWbemObject.Properties_(strPropertyName)/ _
    (1024^3),1)
		If Not IsNull(Gb) Then
			strDiskSize= "Disk Size = " & Gb & " Gb"
		Else
			strDiskSize= "Disk Size = Not Available"
		End If
		If Err <> 0 Then
			strDiskSize= "Disk Size = Not Available"
			Err.Clear
			Err.Number = 0
		End If
		GetInventory = GetInventory & vbNewLine &  strDiskSize
	Next


'  Get the Username  
	strClassName = "Win32_ComputerSystem"
	strPropertyName = "UserName"
	Set objWbemObjectSet = objWbem.ExecQuery("Select " & _
  strPropertyName & " From " & _ 
		strClassName)
	For Each objWbemObject In objWbemObjectSet
		strLastUser="User Last = " & _
    objWbemObject.Properties_(strPropertyName)
	    If Err <> 0 Then
         strLastUser= "User Last = Not Available"
         Err.Clear
         Err.Number = 0
      End If
		GetInventory = GetInventory & vbNewLine &  strLastUser
	Next

	strPropertyName = "DomainRole"
	Set objWbemObjectSet = objWbem.ExecQuery("Select " & _
  strPropertyName & " From " & _ 
		strClassName)
	For Each objWbemObject In objWbemObjectSet
       strMsg = "Domain Role = "
       objRole=objWbemObject.Properties_(strPropertyName)
      Select Case objRole
      Case 0 
        strRole="Standalone Workstation"
        WriteInv strSaveDir & "Investigate_These_Hosts.txt", _
        strComputer 
      Case 1
        strRole="Member Workstation"
      Case 2
        strRole="Standalone Server"
        WriteInv strSaveDir & "Investigate_These_Hosts.txt", _
        strComputer
      Case 3
        strRole="Member Server"
      Case 4
        strRole="Backup Domain Controller"
      Case 5
        strRole="Primary Domain Controller"
      Case Else
        strRole="Unknown Type"
      End Select
      strMsg=strMsg & strRole
      GetInventory = GetInventory & vbNewLine & strMsg

  Next
If Len(GetInventory)=0 Then
	WriteInv strSaveDir & "UpdateWMI.txt", strComputerName
End if

Set objWbemObjectSet = _
GetObject("winmgmts://" & strComputer).ExecQuery _
("select FreeSpace,Size,Name from Win32_LogicalDisk where DriveType=3")
diskInfo = vbNewLine & "Disk Report - "	& vbNewLine
For Each instance In objWbemObjectSet 
	diskName = instance.name
	diskSize = instance.size
	diskFree = instance.FreeSpace
	diskinfo = diskinfo & diskName  & " [Capacity: " & _
  round(diskSize/1024/1024/1024) & "GB;  Available: " & _
  round(diskFree/1024/1024/1024)& "GB]" & vbNewLine
Next

Set objWbemObjectSet = GetObject("winmgmts://" & _
strComputer).ExecQuery("select * from Win32_Share")
shareInfo = vbNewLine & "System Shares Report - "	& vbNewLine
For Each instance In objWbemObjectSet
	shareName = instance.name & vbtab
	sharePath = instance.path & vbtab
	shareDesc = instance.description
	shareInfo = shareInfo & shareName & sharePath & shareDesc & _
  vbcrlf
Next
GetInventory = GetInventory & vbNewLine & diskinfo & _
vbNewLine & shareinfo

Set objWbem = Nothing

End Function
'====================================================================
' Usage Subroutine
Sub Usage()
   wscript.echo "Usage:" & vbNewLine &_
        "C:\> wscript | cscript  wbem.vbs  [Hostname]" & _
        vbNewLine &_
        "Hostname: Optional target host. Local host if omitted."
   WScript.Quit(0)
End Sub

'====================================================================
Function reachable(HostName)

Dim wshShell, fso, tfolder, tname, TempFile, results, retString, ts
Const ForReading = 1, TemporaryFolder = 2
reachable = False
Set wshShell = CreateObject("wscript.shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set tfolder = fso.GetSpecialFolder(TemporaryFolder)
tname = fso.GetTempName
TempFile =fso.buildpath(tfolder, tname)
'-w 100000 is 5 mins worth of timeout
'"cmd /c ping -n 2 -w 500 " & HostName & ">" & TempFile, 0, True
wshShell.Run "%COMSPEC% /c ping.exe -n 2 -w 500 " & HostName & _
">" & TempFile, 0, True
Set results = fso.GetFile(TempFile)
Set ts = results.OpenAsTextStream(ForReading)
Do While ts.AtEndOfStream <> True
    retString = ts.ReadLine
    If InStr(retString, "Reply") > 0 Then
        retString = GetIPAddress(retString)
        reachable = retString
        'LogAction retString
        'LogAction retString & vbTab, LogCompAccts
        Exit Do
    End If
Loop
If InStr(1, retString, ".") Then
    reachable = retString 'Left(retString,Len(RetString))
Else
    reachable = ""
End If

ts.Close
results.Delete
End Function
'====================================================================
Function GetIPAddress(strIPResponse)

strColon = InStr(1, strIPResponse, ":")
strRealAddress = Left(strIPResponse, strColon - 1)
strStart = InStrRev(strRealAddress, "m")
strEnd = Len(strRealAddress) - (strStart + 1)
GetIPAddress = Right(strRealAddress, strEnd + 1)
GetIPAddress=Trim(CStr(GetIPAddress))

End Function
'====================================================================
Function GetSID(HostName)

Dim wshShell, fso, tfolder, tname, TempFile, results, retString, ts
Const ForReading = 1, TemporaryFolder = 2

Set wshShell = CreateObject("wscript.shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set tfolder = fso.GetSpecialFolder(TemporaryFolder)
tname = fso.GetTempName
TempFile =fso.buildpath(tfolder, tname)
strCommand=fso.BuildPath(strScriptPath, "psgetsid.exe")
GetSID ="SID = Not Available"

strBlast="cmd /c " & Chr(34) & strCommand  & Chr(34) & " \\" & _
HostName & " >" &  TempFile 

wshShell.Run strBlast , 0, True

Set results = fso.GetFile(TempFile)
Set ts = results.OpenAsTextStream(ForReading)
Do While ts.AtEndOfStream <> True
    retString = ts.ReadLine
    If InStr(retString, "S-") > 0 Then
        GetSID="SID = " & Right(retString,Len(retString)-1)
		'retString = GetSIDAddress(retString)
        Exit Do
    End If
Loop

ts.Close
results.Delete

End Function
'====================================================================
Function GetInstalledAppsReg(HostName)

On Error Resume Next

Dim oRegistry, sBaseKey, iRC, sKey, arSubKeys, sValue

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE

Set oRegistry = GetObject("winmgmts:\\" & HostName & _
"/root/default:StdRegProv")

sBaseKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
 
iRC = oRegistry.EnumKey(HKLM, sBaseKey, arSubKeys)

For Each sKey In arSubKeys
  iRC = oRegistry.GetStringValue(HKLM, sBaseKey & sKey, _
  "DisplayName", sValue)

  If iRC <> 0 Then
    oRegistry.GetStringValue HKLM, sBaseKey & sKey, _
    "QuietDisplayName", sValue
  End If

  If sValue <> "" Then
    If GetInstalledAppsReg = "" Then
    	GetInstalledAppsReg = sValue & vbNewLine
    Else
      GetInstalledAppsReg = GetInstalledAppsReg & _
      sValue & vbNewLine
    End If
  ElseIf Err <> 0 Then
	  GetInstalledAppsReg =  "Installed App Name Not Available"
  	err.clear
	  err.Number=0
  End If 
Next
 
End Function
 
'In addition you can also use the Windows Installer object
'indirectly with WMI (the Win32_Product Class) to get a list
'of software that is installed with MSI, but this list will
'most likely only be a subset of the Uninstall listing anyway,
'so it 's really no point.
'
'Here is a WMI script that lists all software installed by MSI:
'====================================================================
Function GetInstalledAppsMSI(HostName)
 
Dim strClassName, strPropertyName, objWbem
Dim objWbemObjectSet, objWbemObject
strClassName = "Win32_Product"
strPropertyName = "Caption"

Set objWbem = GetObject("winmgmts:\\" & HostName)
Set objWbemObjectSet = objWbem.ExecQuery( _
"Select Name, Version From " & strClassName)

For Each objWbemObject In objWbemObjectSet
  Debug.Print HostName & " " & objWbemObject.Properties_("Name")
  Debug.Print HostName & " " & "  Version: " & _
  objWbemObject.Properties_("Version")
Next
 
End Function
'====================================================================
Function ProcessProcesses (strComputer)

Dim Procs, Proc, strSysState, DateTime

On Error Resume Next

Set Procs = GetObject _
("winmgmts:{ImpersonationLevel = Impersonate}\\" & _
		strComputer).InstancesOf("Win32_Process")

strSysState = ""

For Each Proc in Procs
  If strSysState <> "" Then
    strSysState = strSysState & vbNewLine & Proc.Caption
    strSysState = strSysState & vbTab & Proc.ProcessID
    strSysState = strSysState & vbTab & Proc.WorkingSetSize
    strSysState = strSysState & vbTab & Proc.PageFileUsage
  Else
    strSysState = Proc.Caption
  End If
Next

ProcessProcesses= strSysState

On Error GoTo 0

End Function
'====================================================================
Function PadNull(strInput)

If Len(strInput) < 1 Then
	PadNull = "0"
Else
	PadNull = strInput
End If

End Function
'====================================================================
Function ProcessServices (strComputer)

Dim Procs, Proc, strSysState, DateTime

On Error Resume Next

Set Procs = GetObject _
("winmgmts:{ImpersonationLevel = Impersonate}\\" & _
		strComputer).InstancesOf("Win32_Service")

strSysState = ""

For Each Proc in Procs			 
	If Proc.Started Then
		'strSysState=strSysState & Proc.Caption & "-" & _
    'Proc.Name & vbNewLine
    strSysState=strSysState & Proc.DisplayName & _
    "-" & Proc.Name & vbNewLine
	End If
Next
ProcessServices = strSysState & vbNewLine

On Error GoTo 0

End Function

'====================================================================
Function GetNicsFinal(strComputer)

On Error Resume Next
'Get a connection to the WMI NetAdapterConfig object
Set NIC1=GetObject("winmgmts:{ImpersonationLevel=Impersonate}\\" & _
strComputer).InstancesOf("Win32_NetworkAdapterConfiguration")
If Err <> 0 Then
  'write to Investigate_These_Hosts.txt
  WriteInv strSaveDir & "Investigate_These_Hosts.txt", strComputer
  GetNicsFinal="Investigate"
  Exit Function
End If  

GetNicsFinal = ""
'For Each of the NICs in the connection
For Each Nic in NIC1
	 'Get the Adapter Description	
     GetNicsFinal = GetNicsFinal & vbNewLine & cstrSeparator 
     GetNicsFinal = GetNicsFinal & vbNewLine & Nic.Description
	 'If IP is enabled on the NIC then let's find out about the NIC
     IF Nic.IPEnabled THEN
		lngCount=UBound(Nic.IPAddress) 
		For i=0 to lngCount
			If i >= 0 Then
				GetNicsFinal = GetNicsFinal & vbNewLine & StrNic & vbNewLine
		        StrIP = Nic.IPAddress(i)
				If StrIP <> "" Then
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "IP Address = " & _
					StrIP
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "MAC Address = " & _
					Nic.MACAddress
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "NIC Service (Short) Name = " & _
					Nic.ServiceName

					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "IP Subnet(s): "
					For j = 0 to UBound(Nic.IPSubnet)
						GetNicsFinal = GetNicsFinal & vbNewLine &  _
            vbTab & Nic.IPSubnet(j)
					Next
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "Internet Database Files Path = " &  _
					Nic.DatabasePath
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "Dead Gateway Detection = " & _
					Nic.DeadGWDetectEnabled
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "IP Gateway(s): "
					For j=LBound(Nic.DefaultIPGateway) to _
          UBound(Nic.DefaultIPGateway)
						GetNicsFinal = GetNicsFinal & vbNewLine &  _
            vbTab & Nic.DefaultIPGateway(j)
					Next
					
					If Nic.DHCPEnabled Then
						GetNicsFinal = GetNicsFinal & vbNewLine &  _
            "DHCP Assigned IP address = " & _
						Nic.DHCPEnabled
						
						GetNicsFinal = GetNicsFinal & vbNewLine &  _
            "DHCP Server = " & _
						Nic.DHCPServer
					End If
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "DNS for WINS Resolution Enabled = " & _
					Nic.DNSEnabledforWINSResolution
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "DNS Host Name = " & _
					Nic.DNSHostName
          If (Nic.DNSHostName <> "") AND _
          (UCase(Nic.DNSHostName) <> UCase(strComputer)) Then
            GetNicsFinal = GetNicsFinal & vbNewLine & _
            " changed output file from " & _
              strSaveDir & strComputer & " to "
            strComputerName=Nic.DNSHostName
            LogFile = strSaveDir & strComputer
            GetNicsFinal = GetNicsFinal & LogFile
          End If

          If fso.FileExists(LogFile) Then
            fso.DeleteFile(LogFile)
          End If
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "DNS Servers:"
					For j=0 to UBound(Nic.DNSServerSearchOrder)
						GetNicsFinal = GetNicsFinal & vbNewLine &  _
            vbTab & Nic.DNSServerSearchOrder(j)
					Next
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "IP Port Filtering Enabled = " & _
					Nic.IPFilterSecurityEnabled
					
					If Nic.IPFilterSecurityEnabled Then
						GetNicsFinal = GetNicsFinal & vbNewLine &  _
            "IP Filtering Enabled."
						
						If Nic.IPSecPermitIPProtocols <> 0 Then
							For j=0 to UBound(Nic.IPSecPermitIPProtocols)
								GetNicsFinal = GetNicsFinal & vbNewLine &  _
                vbTab & "Protocol: " & _
								Nic.IPSecPermitIPProtocols(j)
							Next
						Else
							GetNicsFinal = GetNicsFinal & vbNewLine &  _
              vbTab & "No Protocols Filtered"
						End If
						
						If Nic.IPSecPermitTCPPorts <> 0 Then
							For j=0 to UBound(Nic.IPSecPermitTCPPorts)
								GetNicsFinal = GetNicsFinal & vbNewLine &  vbTab & _
                "TCP Port: " & _
								Nic.IPSecPermitTCPPorts(j)
							Next
						Else
							GetNicsFinal = GetNicsFinal & vbNewLine &  _
              vbTab & "No TCP Ports Filtered"
						End If
						
						If Nic.IPSecPermitUDPPorts <> 0 Then
							For j=0 to UBound(Nic.IPSecPermitUDPPorts)
								GetNicsFinal = GetNicsFinal & vbNewLine &  _
                vbTab & "UDP Port: " & _
								Nic.IPSecPermitUDPPorts(j)
							Next
						Else
							GetNicsFinal = GetNicsFinal & vbNewLine &  _
              vbTab & "No UDP Ports Filtered"
						End If

					End If
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "LMHOSTS Lookup Enabled = " & _
					Nic.WINSEnableLMHostsLookup	
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "WINS Lookup File = " & _
					Nic.WINSHostLookupFile
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "Primary WINS Server = " & _
					Nic.WINSPrimaryServer
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "Secondary WINS Server = " & _
					Nic.WINSSecondaryServer
					
					GetNicsFinal = GetNicsFinal & vbNewLine &  _
          "WINS Scope ID = " & Nic.WINSScopeID

				End If
			End If
		Next
     END IF
Next

GetNicsFinal = GetNicsFinal & vbNewLine &   _
"==================================" 

GetNicsFinal = GetNicsFinal & vbNewLine &  _
"Net Link Speed Report:" 
Set oNDIS = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate, (Security)}!\\" & _
strComputer & "\root\wmi").InstancesOf("MSNDIS_LinkSpeed")

For Each Obj In oNDIS
 GetNicsFinal = GetNicsFinal & vbNewLine & _
 Obj.InstanceName
 GetNicsFinal = GetNicsFinal & vbNewLine & _
 Obj.NDISLinkSpeed / 10 & "Kbps"

 If (Obj.Active = True) Then
  GetNicsFinal = GetNicsFinal & vbNewLine & _
  "Active:  Yes"
 Else
  GetNicsFinal = GetNicsFinal & vbNewLine & _
  "Active:  No"
 End If

 
Next

GetNicsFinal = GetNicsFinal & vbNewLine & _
"---------------------------------"

End Function
'=======================================================

Sub LogAction (strEntry)

Dim strErrMsg, f

On Error Resume Next

set f = fso.OpenTextFile(LogFile, ForAppending, True, -2)
f.WriteLine strEntry
If Err <> 0 Then
  Err.Clear
  randomize
  Delay=Right(FormatNumber(rnd,2),2)
  wscript.sleep(Delay)
  LogAction strEntry
End If
f.close

wscript.echo strEntry

On Error Goto 0

End Sub

'====================================================================

Sub CheckScriptHost()

If InStr(LCase(wscript.fullname),"cscript") = 0 Then
	strMsg = _
  "Script must be run by CScript.exe. Terminating this " & _
	"script, changing the default script engine and restarting" & _
	" execution."
	Dim objShell
	Set objShell=CreateObject("wscript.shell")
	strExec = "cscript.exe //NoLogo //H:cscript //S"
	objShell.Run strExec,0,TRUE
        
	strExec ="cmd /c " & Chr(34) & "cscript.exe " & Chr(34) & _
	wscript.scriptfullname & Chr(34)& Chr(34)
        
	objShell.Run strExec,,False
	Wscript.Quit
End If

End Sub

'====================================================================

Function ExecutingFrom()

Dim strScriptPath

strScriptPath=Left(wscript.scriptfullname, _
Len(wscript.scriptfullname)-Len(wscript.scriptname))

If Right(strScriptPath,1) <> "\" Then
	strScriptPath=strScriptPath & "\"
End If

ExecutingFrom=strScriptPath

End Function

'====================================================================

Sub MakeExist (strFolderPath)

On Error Resume Next

If IsEmpty(strFolderPath) Then Exit Sub

LogBuffer "MakeExist:strFolderPath = " & strFolderPath, 4
If NOT (fso.FolderExists(strFolderPath)) Then
	Set f = fso.CreateFolder(strFolderPath)
	If Err <> 0 Then
    strMsg = "Create " & strFolderPath & " = Fail"
    'LogBuffer strMsg,1
    'LogBuffer err.number & ", " & err.Description,1
    Err.Clear
	Else
    strMsg = "Create " & strFolderPath & " = Success"
    'LogBuffer strMsg,2
	End If
End If

On Error GoTo 0

End Sub

'====================================================================

Sub CheckDestDirs(strFolder)

On Error Resume Next

LogAction "Checking " & strFolder

If IsEmpty(strFolder) Then Exit Sub
If NOT (fso.FolderExists(strFolder)) Then
	strPath=Split(strFolder,"\")
	If NOT Instr(strPath(0),":") > 0 Then
    BasePath ="\\" & strPath(0) & strPath(1)
    For x = 2 to UBound(strPath)
    	BasePath=BasePath & "\" & strPath(x)
    	MakeExist(BasePath)
    Next
	Else
    BasePath=strPath(0)
    For x=1 to UBound(strPath)
    	BasePath=BasePath & "\" & strPath(x)
    	MakeExist (BasePath)
    Next
	End If
End If

On Error GoTo 0

End Sub
'====================================================================
Sub WriteInv (strFile,strValue)

On Error Resume Next

'wscript.echo strFile & strValue
Set OutFile=fso.OpenTextFile(strFile,ForAppending,True,False)
OutFile.WriteLine strValue
If Err <>0 Then
  Err.Clear
    randomize
  Delay=Right(FormatNumber(rnd,2),2)
  wscript.sleep(Delay)
  WriteInv strFile,strValue
End If  

Outfile.Close

wscript.echo strFile & ": " & strValue

Set OutFile=Nothing

On Error Goto 0

End Sub
'====================================================================
Function WMIGetSID(strComputer)

Set SIDs=GetObject("winmgmts:{ImpersonationLevel=Impersonate}\\" & _
strComputer).InstancesOf("Win32_Account")

For Each SID in SIDS
  If SID.SIDType=1 Then
    WMIGetSID = BaseSID(SID.SID)
    Exit Function
  End if
Next

End Function
'====================================================================
Function BaseSID(strSID)

arrSplit=Split(strSID, "-")

For i= 0 to UBound(arrSplit)-1
  If BaseSID <> "" Then
    BaseSID=BaseSID & "-" & arrSplit(i)
  Else 
    BaseSID=arrSplit(i)
  End If
Next

End Function
'====================================================================
Function ConvWbemTime(IntervalFormat)
  Dim sYear, sMonth, sDay, sHour, sMinutes, sSeconds
  sYear = mid(IntervalFormat, 1, 4)
  sMonth = mid(IntervalFormat, 5, 2)
  sDay = mid(IntervalFormat, 7, 2)
  sHour = mid(IntervalFormat, 9, 2)
  sMinutes = mid(IntervalFormat, 11, 2)
  sSeconds = mid(IntervalFormat, 13, 2)

  ' Returning format yyyy-mm-dd hh:mm:ss
  ConvWbemTime = sYear & "-" & sMonth & "-" & sDay & " " _
               & sHour & ":" & sMinutes & ":" & sSeconds
End Function
