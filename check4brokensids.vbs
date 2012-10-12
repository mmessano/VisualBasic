 'check4brokensids.vbs ver 1.3

'4-6-03 Alan Kaplan for VA VISN6 alan.kaplan@med.va.gov, alan@akaplan.com
'This script searchs for unresolved SIDS on a level of folders iterated below
'a specified starting point. It was written to identify home directories where
'the user was deleted
'The results are logged to a CSV file on your desktop.
'Requires that home directory server has WMI. This would include Win2k and later
'or NT 4 with WMICore installed by SMS or manually.
'4-8-03 added folder size, removed logging of "okay" folders
'9-5-03 Fixed crash issue ver 1.2
'9-8 ver 1.3 fixed file size always being checked, added error handling for MyInfo

Option Explicit
Dim strStartFolder    'target file or folder path
Dim oADSSecurity    
Dim fso,rserver,f
Dim objSubFolder, objFolder,userpath
Dim wshShell, iFolder, message, quitmessage
dim quote, strUNCStart, writetype
Dim oACE, oTargetSD, oDACL, strFsize, strCheckFldr, strUNCFolder        
dim broken, retval,status, objLocator, objService
dim strCheckSize
Dim logpath, logfile, appendout, wmiFileSecuritySetting
Set fso = CreateObject("Scripting.FileSystemObject")
set wshShell = WScript.CreateObject ("WScript.Shell")

broken = False
quote=chr(34)

syscheck

If (Not IsCScript()) Then         'If not CScript, re-run with cscript...
    WshShell.Run "CScript.exe " & quote & WScript.ScriptFullName & quote, 1, true
WScript.Quit     '...and stop running as WScript
End If
On Error goto 0

'These are over written if user running has %HOMESHARE% set
rServer=wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
userpath = "d:\Home"

message = "This script will search for unresolved SIDS on a level of folders below a specified starting point. "
message = message & "It was written to identify home directories where the user was deleted. "
message = message & "The queried server must support WMI -- Windows 2000 or later or NT4 with WMI installed. " & vbcrlf & vbcrlf
message = message & "File Size can be optionally reported. "
message = message & "The WMI interface is not very fast, and we are doing an interative query of each Trustee. So be patient!"& vbcrlf & vbcrlf
message = message & "The folders with broken SIDs are logged to a CSV file on your desktop."
retval=MsgBox(message, vbokcancel,"Welcome")
If retval = vbcancel Then
    quitmessage = "Quitting at your request."
    Abort
End If

MyInfo

message = "Enter the name of local system or remote server without"&_
    " leading " & quote & "\\" & quote
rServer= InputBox(message,"Server Name",rserver)
rserver = UCase(rserver)

If rServer = "" Then
    quitmessage = "Quitting at your request."
    Abort
End If

message = "What is the starting local path to start searching?" & vbcrlf & vbcrlf
strStartFolder= InputBox(message,"Start Path",userpath)

If strStartfolder = "" Then
    quitmessage = "Quitting at your request."
    Abort
End If

message = "Do you want the folder size reported for folders found with broken SID?" &_
    " This will add some time for each broken folder." & vbcrlf & vbcrlf
retval= MsgBox(message,vbyesno,"Check Size")

If strCheckSize = vbyes Then
    strCheckSize = True
Else
    strCheckSize = False
End If
    
logsetup
WMIConnect

'convert to UNC
strUNCStart = "\\"&rserver&"\"& Replace(strStartFolder,":","$")

If not fso.FolderExists(strUNCStart) Then
    quitmessage = strUNCStart & " Not Found"
        abort
End If


Set objFolder = fso.getfolder(strUNCStart)
Set objSubFolder = objFolder.SubFolders

for Each iFolder in objSubFolder
    broken = False
    strCheckFldr = iFolder.name
    strUNCFolder = strUNCStart & "\"& strCheckFldr
    DisplayACLs strCheckFldr
    
If (strCheckSize And broken) Then
    strFSize = ShowFolderSizeMB(strUNCFolder)
Else
    strFsize = "Not Checked"
End If

    If broken = False Then
        'If you want to log success and failure,
        'comment following line, and uncomment the one below
        WScript.Echo strUNCFolder & " skipped -- Okay"
         'echoandlog strUNCFolder & ",skipped -- Okay"
        Else
    echoandlog strUNCFolder & ","& strFsize
    End If
Next

appendout.Close

If (MsgBox("Script completed successfully." & vbCrLf & "Open log "& quote& logfile & quote& " ?",_
    vbYesNo + vbQuestion,"Script Complete") = vbYes) Then
    WshShell.Run(quote&logfile& quote)
End If

Set wmiFileSecuritySetting = nothing
set fso = nothing
set objfolder = nothing
set objsubfolder = nothing

Wscript.Quit        'Script ends

Sub WMIConnect()
'Connect once, query many!
On Error Resume next
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objService = objLocator.ConnectServer(rserver, "root\cimv2")', USERID, PASSWORD)
objService.Security_.ImpersonationLevel = 3
If Err <> 0 Then
    quitmessage = "Fatal error try to connect to WMI on " & rserver & "." & vbcrlf & vbcrlf &_
    "This could be because the remote server is NT 4 without WMI installed"&_
    ", or because of an authentication problem."
    Abort
End If
On Error goto 0

End Sub


'*********** Functions and subs

Sub DisplayACLs(strNextFolder)
Dim filespec, wmiDacl, i, wmisd
filespec= strStartFolder & "\" & strnextfolder
filespec = Replace(filespec, "\", "\\")

on error resume next
Set wmiFileSecuritySetting = objService.Get("Win32_LogicalFileSecuritySetting.Path='" &filespec & "'")
'WScript.Echo Err.Description
'This WMIDACL part of this subroutine is excerpted from ACLS.VBS by Marcin Policht
'May 13, 2002 Scripting NTFS permisions with ADSI (Part 2)
'http://www.serverwatch.com/tutorials/article.php/1476741

' Get the security descriptor, and store it in wmiSD.
retval = wmiFileSecuritySetting.GetSecurityDescriptor(wmiSD)
'WScript.Echo Err.Description

' Retrieve the information from the security descriptor.
Set wmiDacl = wmiSD.Properties_.Item("Dacl")
'WScript.Echo Err.Description


For i = 0 To UBound(wmiDacl.Value)
If broken = True Then Exit For
    retval = wmiDacl.Value(i).Properties_.Item("Trustee").Value.Properties_.Item("Name")
    If IsNull(retval) Then    'unresolved SIDS show as a NULL entry
        Broken = True
    End If
next
on error goto 0
End sub

sub syscheck()
    Dim major,minor, ver, key, key2
    on error resume next    
    Major = (ScriptEngineMinorVersion())
    Minor = (ScriptEngineMinorVersion())/10
    Ver = major + minor
    'Need version 5.5
        If err.number or ver < 5.5 then
        quitmessage = "You have WScript Version " & ver & ". Please load Version 5.5"
    End If
    
    'Test for ADSI
    err.clear
    key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\{E92B03AB-B707-11d2-9CBD-0000F87A369E}\version"
    key2 = WshShell.RegRead (key)
    if err <> 0 then
         quitmessage = quitmessage & "ADSI must be installed on local workstation to continue" & vbCrLf
        abort
        End if    
End Sub

Function Abort()    'error message handler
    WshShell.Popup quitmessage,0,"Abort",vbCritical
    WScript.Quit
End Function

Function IsCScript()
' Check whether CScript.exe is the host.
If (InStr(UCase(WScript.FullName), "CSCRIPT") <> 0) Then
IsCScript = True
Else
IsCScript = False
End If
End Function

Sub logsetup()
    'Setting up log
    On Error goto 0
    Dim arTemp
    logpath= wshShell.SpecialFolders("Desktop") & "\"
    arTemp= Split(WScript.ScriptName,".")    'Script Name found"
    logfile = logpath & artemp(0)& ".csv"    'append .CSV
    'setup Log
    writeType = 2 ' forwriting                'presume for writing. Usually done as constant...

If fso.FileExists(logfile) Then
retval = MsgBox("Logfile Exists, do you want to append?",vbyesno + vbdefaultbutton1,"Old Log File")    
    If retval = vbyes Then
        writeType = 8                        'change type to append
    End If
End If
    On Error Resume next
    set AppendOut = fso.OpenTextFile(logfile, Writetype, True)
    If Err <> 0 Then
        MsgBox "You must close the log file!",vbcritical + vbinformation,"Fatal Error"
        WScript.Quit
    End If
    On Error goto 0
    If writetype = 2 Then             'only write header if new file
        appendout.writeline "UNC to Folder with Broken SID,Size(MB)"
    End If
End sub


Sub EchoAndLog (message)
'Echo output and write to log
    Wscript.Echo message
    AppendOut.WriteLine message
End Sub

Sub MyInfo()
on error resume next
'Some unnecessary flash to seed the default fields for prompts
'Find user's home directory
    Dim sharename,rlanmanobj, rshobj, mymane, strhomedir,arr
    strhomedir = wshShell.ExpandEnvironmentStrings("%homeshare%")
    If strhomedir = "" Then Exit sub
    arr = Split(strHomedir, "\")        'create array of name split by \
    rServer = arr(2)                    'get home directory server and share
    sharename = arr(3)
    'get physcal path
    On Error Resume Next
    Set RLanManObj = GetObject("WinNT://"& rserver &"/LanmanServer")
    Set RShObj = RLanManObj.GetObject("Fileshare",sharename)
    If Err <> 0 Then Exit Sub
    userpath=RShObj.path
    userpath = Left(userpath,InStrRev(userpath, "\")-1)    'go up one level
    On Error goto 0
on error goto 0
End Sub

Function ShowFolderSizeMB(filespec)
Set f = fso.GetFolder(filespec)
ShowFolderSizeMB = Round((f.size/1048576),2)
End Function

Keywords: Sid, Home Directory,orphaned, Fso, Wmi




Posted by: Alan Kaplan     
Date: 2/25/2004 9:15:12 PM
Comment: I got a good suggestion for error trapping from Andy Ray:


Function ShowFolderSizeMB(filespec)

On Error Resume next
Set f = fso.GetFolder(filespec)
ShowFolderSizeMB = Round((f.size/1048576),2)

If Err.Number <> 0 Then
If Err.Number = 70 Then

ShowFolderSizeMB = "Error: " & Err.Description
Else
ShowFolderSizeMB = "Error: " & Err.Description & ": " & Err.Number
End If
End If
Err.Clear

End Function

Posted by: Mr Paul     
Date: 10/11/2004 4:48:25 AM
Comment: Nice script, but I had to modify it to get the folder size reporting feature to work.

The following code:
retval= MsgBox(message,vbyesno,"Check Size")

If strCheckSize = vbyes Then
strCheckSize = True
Else
strCheckSize = False
End If

Should be:
retval= MsgBox(message,vbyesno,"Check Size")

If retval = vbyes Then
strCheckSize = True
Else
strCheckSize = False
End If
