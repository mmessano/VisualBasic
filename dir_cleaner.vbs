 Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' AUTHOR : Alain AUCORDIER
' CREATED : 11/12/2004
' FUNCTION: Perform Directory cleaning by Age of Files.
' MODIFIED: None
'------------------

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Constants Declaration
'------------------

Const ForReading = 1 , ForWriting = 2 , ForAppending = 8
'
' Define a ADS_RIGHTS_ENUM constants:
'
const ADS_RIGHT_DELETE = &h10000
const ADS_RIGHT_READ_CONTROL = &h20000
const ADS_RIGHT_WRITE_DAC = &h40000
const ADS_RIGHT_WRITE_OWNER = &h80000
const ADS_RIGHT_SYNCHRONIZE = &h100000
const ADS_RIGHT_ACCESS_SYSTEM_SECURITY = &h1000000
const ADS_RIGHT_GENERIC_READ = &h80000000
const ADS_RIGHT_GENERIC_WRITE = &h40000000
const ADS_RIGHT_GENERIC_EXECUTE = &h20000000
const ADS_RIGHT_GENERIC_ALL = &h10000000
const ADS_RIGHT_DS_CREATE_CHILD = &h1
const ADS_RIGHT_DS_DELETE_CHILD = &h2
const ADS_RIGHT_ACTRL_DS_LIST = &h4
const ADS_RIGHT_DS_SELF = &h8
const ADS_RIGHT_DS_READ_PROP = &h10
const ADS_RIGHT_DS_WRITE_PROP = &h20
const ADS_RIGHT_DS_DELETE_TREE = &h40
const ADS_RIGHT_DS_LIST_OBJECT = &h80
const ADS_RIGHT_DS_CONTROL_ACCESS = &h100
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' ADS_ACETYPE_ENUM
' Ace Type definitions
'
const ADS_ACETYPE_ACCESS_ALLOWED = 0
const ADS_ACETYPE_ACCESS_DENIED = &h1
const ADS_ACETYPE_SYSTEM_AUDIT = &h2
const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &h5
const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &h6
const ADS_ACETYPE_SYSTEM_AUDIT_OBJECT = &h7
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' ADS_ACEFLAGS_ENUM
' Ace Flag Constants
'
const ADS_ACEFLAG_UNKNOWN = &h1
const ADS_ACEFLAG_INHERIT_ACE = &h2
const ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE = &h4
const ADS_ACEFLAG_INHERIT_ONLY_ACE = &h8
const ADS_ACEFLAG_INHERITED_ACE = &h10
const ADS_ACEFLAG_VALID_INHERIT_FLAGS = &h1f
const ADS_ACEFLAG_SUCCESSFUL_ACCESS = &h40
const ADS_ACEFLAG_FAILED_ACCESS = &h80

Const ADS_RIGHT_LIST = &H4
Const ADS_RIGHT_READ = &H80000000
Const ADS_RIGHT_EXECUTE = &H20000000
Const ADS_RIGHT_WRITE = &H40000000
' Const ADS_RIGHT_DELETE = &H10000
Const ADS_RIGHT_FULL = 2032127 ' &H10000000
Const ADS_RIGHT_CHANGE_PERMS = &H40000
Const ADS_RIGHT_TAKE_OWNERSHIP = &H80000
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Global Variables
'------------------
dim oAdsSecurity,oFS, oDebug, oOutput, tExceptions(100), nExceptions
Dim oArchive, oRestore
Dim scanPath        ' Path to the top of Directory tree to Scan
Dim scanDrive        ' The Drive related to scanPath
Dim scanServer        ' The Server related to scanPath (if not Local)
Dim scanShare        ' The Share related to scanPath (if not Local)
Dim accountDomain    ' Account Domain Name
Dim pastMonths        ' Number of Months Old we track
Dim pastDays        ' Number of Days Old we track
Dim limitDate        ' Limit Date for Archiving
Dim logPath        ' Directory Path to logfiles
Dim archivePath        ' Directory Path for Archiving Old Directories
Dim currentDrive
Dim currentPath
Dim gSize
'
' Cache Users & Group to speedup access
'
Dim nCachedNames
Dim nameCache
'
' Directory Stack
'
Dim directoryStack(100)
Dim stackDepth
Dim fDebug
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Start Program
'------------------
On Error Resume Next
    fDebug = False
    fDebug = True
    nCachedNames = 0
    Call Main()
    WScript.Quit(0)
Sub Main()
Dim Args
Dim dRoot
    On Error Resume Next
    set Args = WScript.Arguments
    Call ProcessCommandLine( Args )
    Call InitProgram()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Start Processing the Root Folder
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dRoot = ProcessFolder ( scanPath )
    Do While Not IsEmptyStack()
        Call ArchiveFolder( PopDirectory() )
    Loop
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Load Results into Excel Spreadsheet
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call ProduceExcelResults()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Mail Results to Administrators
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call SendResults()
    Call TerminateProgram()
    WScript.Quit(0)
End Sub
Sub TerminateProgram()
    oArchive.WriteLine currentDrive
    oArchive.WriteLine "CD " & Chr(34) & currentPath & Chr(34)
    oOutput.Close
    oDebug.Close
    oArchive.Close
    oRestore.Close
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Perform All Initializations, create mandatory objets, logfiles, set counters, etc ...
'------------------

Sub InitProgram()
Dim scriptName , scriptDir
    stackDepth = 0
    scriptName = WScript.ScriptFullName
    scriptDir = Left(scriptName,InStrRev(scriptName,"\",-1,1)-1)
    currentDrive = Left(scriptName,2)
    currentPath = Right(scriptDir,Len(scriptDir)-2)
    Set oFS = WScript.CreateObject("Scripting.FileSystemObject")
    Set nameCache = WScript.CreateObject("Scripting.Dictionary")
    Set oOutput = oFs.OpenTextFile(logPath & "DIRAGING.TXT",ForWriting,True)
    Set oDebug = oFs.OpenTextFile(logPath & "DIRDEBUG.TXT",ForWriting,True)
    Set oArchive = oFs.OpenTextFile(logPath & "ARCHIVE.BAT",ForWriting,True)
    Set oRestore = oFs.OpenTextFile(logPath & "RESTORE.BAT",ForWriting,True)
    '
    ' Generate all Headers
    '
    ' Log File
    oOutput.WriteLine "Directory Path" & Chr(9) & "LastModified" & Chr(9) & "Size" & Chr(9) & "Owner" & Chr(9) & "Readers"
    ' Archive File
    oArchive.WriteLine "SET SCANPATH=" & Chr(34) & scanPath & Chr(34)
    oArchive.WriteLine "SET SCANDRIVE=" & Chr(34) & scanDrive & Chr(34)
    oArchive.WriteLine "SET SCANSHARE=" & Chr(34) & scanShare & Chr(34)
    oArchive.WriteLine "REM ECHO OFF"
    oArchive.WriteLine "NET USE " & scanDrive & " /D"
    oArchive.WriteLine "NET USE " & scanDrive & " " & scanShare
    oArchive.WriteLine scanDrive
    oArchive.WriteLine "MD " & Chr(34) & archivePath & Chr(34)
    oArchive.WriteLine "CD " & Chr(34) & archivePath & Chr(34)
    ' Restore File
    oRestore.WriteLine "SET SCANPATH=" & Chr(34) & scanPath & Chr(34)
    oRestore.WriteLine "SET SCANDRIVE=" & Chr(34) & scanDrive & Chr(34)
    oRestore.WriteLine "SET SCANSHARE=" & Chr(34) & scanShare & Chr(34)
    oRestore.WriteLine "ECHO OFF"
    oRestore.WriteLine "NET USE " & scanDrive & " /D"
    oRestore.WriteLine "NET USE " & scanDrive & " " & scanShare
    oRestore.WriteLine scanDrive
    oRestore.WriteLine "CD " & scanPath
    Set oAdsSecurity = CreateObject("ADsSecurity")
    Call AddException( "BUILTIN\Administrators" )
    Call AddException( "BUILTIN\Replicator" )
    Call AddException( "CREATOR OWNER" )
    Call AddException( "NT AUTHORITY\SYSTEM" )
    AddCacheName "EveryOne","All Users"
    AddCacheName "BUILTIN\Users","All Users"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' account domain administrators should be excluded as they have access to any directory
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call AddException( accountDomain & "\Administrators" )
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Add a user of group Exception not to be processed
'------------------
Sub AddException( pException )
    tExceptions(nExceptions) = pException
    nExceptions = nExceptions + 1
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Check if a user or Group is an exception not to be processed
'------------------
Function IsException( pTrustee )
Dim I
    IsException = False
    For I=0 To nExceptions - 1
        If tExceptions(I) = pTrustee Then
            IsException = True
            Exit For
        End If
    Next
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Process Script Arguments
'------------------
Sub ProcessCommandLine( pArgs )
On Error Resume Next
    If( pArgs.Count < 1 ) then
        Call DisplayUsage ( )
        WScript.Quit(1)
    End If
    If ( Not pArgs.Named.Exists("LOGPATH") ) Then
        logPath = "C:\TEMP"
    Else
        logPath = pArgs.Named("LOGPATH")
    End If
    If Right(logPath,1) <> "\" Then
        logPath = logPath & "\"
    End If
    If ( Not pArgs.Named.Exists("SCANPATH") ) Then
        scanPath = "C:\"
    Else
        scanPath = pArgs.Named("SCANPATH")
    End If
    scanDrive = Left(scanPath,2)
    scanServer = GetScanServer( scanDrive )
    scanShare = GetScanShare( scanDrive )
    If ( Not pArgs.Named.Exists("ARCHIVEPATH") ) Then
        archivePath = scanDrive & "\Archives"
    Else
        archivePath = pArgs.Named("ARCHIVEPATH")
    End If
    If Right(logPath,1) <> "\" Then
        logPath = logPath & "\"
    End If
    If ( Not pArgs.Named.Exists("DOMAIN") ) Then
        accountDomain = GetDefaultDomain()
    Else
        accountDomain = pArgs.Named("DOMAIN")
    End If
    If ( Not pArgs.Named.Exists("MONTHS") ) Then
        pastMonths = 18
        limitDate = DateAdd("m",-1*pastMonths,Now())
    Else
        pastMonths = CInt(pArgs.Named("MONTHS"))
        limitDate = DateAdd("m",-1*pastMonths,Now())
    End If
    If ( pArgs.Named.Exists("DAYS") ) Then
        pastDays = CInt(pArgs.Named("DAYS"))
        limitDate = DateAdd("d",-1*pastDays,Now())
    End If
    WScript.Echo "SCANPATH : " & scanPath
    WScript.Echo "SCANDRIVE : " & scanDrive
    WScript.Echo "SCANSERVER : " & scanServer
    WScript.Echo "SCANSHARE : " & scanShare
    WScript.Echo "LOGPATH : " & logPath
    WScript.Echo "ARCHIVEPATH : " & archivePath
    WScript.Echo "DOMAIN : " & accountDomain
    WScript.Echo "MONTHS : " & pastMonths
    WScript.Echo "LIMIT : " & CStr(limitDate)
end Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Display Script Usage
'------------------
Sub DisplayUsage
    WScript.Echo "USAGE: cscript.exe dir_cleaner.vbs [OPTIONS]"
    WScript.Echo
    WScript.Echo "Where: OPTIONS are:"
    WScript.Echo vbTab & "/DOMAIN:NameOfAccountDomain (Optional, Default=" & GetDefaultDomain() & ")"
    WScript.Echo vbTab & "/MONTHS:NumberOfMonthsOld (Optional, Default=18)"
    WScript.Echo vbTab & "/LOGPATH:PathToOutputFiles (Optional, Default=C:\TEMP)"
    WScript.Echo vbTab & "/ARCHIVEPATH:PathToArchiveFiles (Optional, Default=\Archives on the scanned Drive)"
    WScript.Echo vbTab & "/SCANPATH:RootPathToScan (Mandatory)"
    WScript.Echo
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Process A Single Folder (Recursively processing SubFolders)
'------------------
Function ProcessFolder ( pPath )
' Recursive folder handling
Dim oFolder,oSubFolders,oSubFolder,oFiles,oFile,dPath
Dim fileDate,subfolderDate,folderDate,dummy
Dim dSize
On Error Resume Next
    dSize = 0
    gSize = 0
    If Right(pPath,1) = "\" Then
        dPath = pPath
    Else
        dPath = pPath & "\"
    End If
    folderDate = CDate("01/01/1970")
    ' Get Folders-collection
    Set oFolder = oFs.GetFolder(pPath)
    Set oSubFolders = oFolder.SubFolders
    Set oFiles = oFolder.files
    folderDate = oFolder.DateLastModified
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Process Files first
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each oFile in oFiles
        fileDate = oFile.DateLastModified
        dSize = dSize + oFile.Size
        
        If fileDate > folderDate Then
            folderDate = fileDate
        End If
    Next
    'WScript.Echo pPath & " : " & CStr(dSize)
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Then Process Sub Folders
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each oSubFolder in oSubFolders
    ' search for subfolders
        subfolderDate = CDate("01/01/1970")
        subfolderDate = oSubFolder.DateLastModified
        subfolderDate = ProcessFolder ( dPath & oSubFolder.Name )
        dSize = dSize + gSize
        'WScript.Echo "GSIZE: " & CStr(gSize)
        'WScript.Echo "DSIZE: " & CStr(dSize)
        If subfolderDate > folderDate Then
            folderDate = subfolderDate
        End If
        If subfolderDate > limitDate Then
            Do While Not IsEmptyStack()
                Call ArchiveFolder( PopDirectory() )
            Loop
        End If
    Next
    gSize = dSize
    ProcessFolder = folderDate
    If folderDate < limitDate Then
        '
        ' Main Folder Ok:
        '    Pop my Subfolders Only !!!!
        '    Push MainFolder
        '
        For Each oSubFolder in oSubFolders
            dummy = PopDirectory()
        Next
        Call PushDirectory (dPath)
        Call ShowFolder( dPath , folderDate , dSize)
        'Call ArchiveFolder( dPath , folderDate )
        'Else
        'WScript.Echo "NEWER: " & pPath & Chr(9) & CStr(folderDate)
    End If
    Set oFolder = Nothing ' release objects
    Set oSubFolders = Nothing
    Set oFiles = Nothing
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Process the Archiving of the Old Folder
'------------------
Sub ShowFolder( pPath , pDate , pSize )
Dim namedOwner, namedReaders
On Error Resume Next
    namedOwner = ""
    namedReaders = ""
    namedOwner = GetNamedOwner( pPath )
    namedReaders = GetNamedReaders( pPath )
    WScript.Echo pPath & Chr(9) & CStr(pDate) & Chr(9) & CStr(pSize) & Chr(9) & namedOwner & Chr(9) & namedReaders
    oOutput.WriteLine pPath & Chr(9) & CStr(pDate) & Chr(9) & CStr(pSize) & Chr(9) & namedOwner & Chr(9) & namedReaders
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Process the Archiving of the Old Folder
'------------------

Sub ArchiveFolder( pPath )
On Error Resume Next
    Call MkdirArchivePath(pPath)
    oArchive.WriteLine "ROBOCOPY " & NoLastSlash(pPath) & " " & BuildArchivePath(pPath) & " /PURGE"
    oRestore.WriteLine "ROBOCOPY" & archivePath & Right(pPath,Len(pPath)-3) & " " & scanPath & " /PURGE"
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get The FullName of Directory Owner
'------------------
Function GetNamedOwner( pPath )
Dim sDirPath
Dim oFileSD,oDacl,myAcl
Dim ownerId,ownerDomain,ownerLogin
On Error Resume Next
    sDirPath = "FILE://" & pPath
    Set oFileSD = Nothing
    Err.Clear
    Set oFileSD = oADsSecurity.GetSecurityDescriptor(CStr(sDirPath))
    If oFileSD Is Nothing Then
        WScript.Echo "ERROR: " & Err.Description
        GetNamedOwner=""
        Exit Function
    End If
    Set oDacl = Nothing
    Err.Clear
    Set oDacl = oFileSD.DiscretionaryACL
    If oDacl Is Nothing Then
        WScript.Echo "ERROR: " & Err.Description
        GetNamedOwner=""
        Exit Function
    End If
    ownerId = oFileSD.Owner
    GetNamedOwner = GetUserById( ownerId )
    Set oDacl = Nothing
    Set oFileSD = Nothing
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get the Name List of Users with access to the Directory (in case the owner is not relevent)
'------------------
Function GetNamedReaders( pPath )
Dim sDirPath
Dim oFileSD,oDacl,readerList,readerTable,namedReader,indexReader,namedReaders
On Error Resume Next
    sDirPath = "FILE://" & pPath
    Set oFileSD = Nothing
    Err.Clear
    Set oFileSD = oADsSecurity.GetSecurityDescriptor(CStr(sDirPath))
    If oFileSD Is Nothing Then
        WScript.Echo "ERROR: " & Err.Description
        GetNamedReaders=""
        Exit Function
    End If
    Set oDacl = Nothing
    Err.Clear
    Set oDacl = oFileSD.DiscretionaryACL
    If oDacl Is Nothing Then
        WScript.Echo "ERROR: " & Err.Description
        GetNamedReaders=""
        Exit Function
    End If
    readerList= GetTrusteeList( oDacl )
    readerTable = Split(readerList,"!",-1,1)
    namedReaders=""
    'For indexReader=0 To UBound(readerTable) - 1
    For indexReader=0 To UBound(readerTable)
        namedReader = GetTrusteeById(readerTable(indexReader))
        If namedReaders = "" Then
            namedReaders = namedReader
        Else
            namedReaders = namedReaders & Chr(9) & namedReader
        End If
    Next
    GetNamedReaders = namedReaders
    Set oDacl = Nothing
    Set oFileSD = Nothing    
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get The Trustee List from the ACL on the Directory
'------------------
Function GetTrusteeList( pDacl )
Dim ace
Dim trusteeTable(500), trusteeCount, indexTrustee
Dim pTrustee, cTrustee
Dim sMask, cMask
On Error Resume Next
    pTrustee = ""
    trusteeCount = 0
    trusteeTable(0) = ""
    For Each ace in pDacl
        cTrustee = ace.Trustee
        cMask = ace.AccessMask
        sMask = GetACLMask( cMask )
        If ( cTrustee <> pTrustee ) Then
            If ( pTrustee <> "" ) Then
                If Not IsException( pTrustee ) Then
                    trusteeTable( trusteeCount ) = pTrustee
                    trusteeCount = trusteeCount + 1
                    If trusteeCount > 10 Then
                        trusteeCount = 10
                        Exit For
                    End If
                End If
            End If
        End If
        pTrustee = cTrustee
    Next
    If Not IsException( cTrustee ) Then
        trusteeTable( trusteeCount ) = pTrustee
        trusteeCount = trusteeCount + 1
        ' Limit number of Trustees
        If trusteeCount > 10 Then
            trusteeCount = 10
        End If
    End If
    GetTrusteeList = trusteeTable(0)
    For indexTrustee = 1 To trusteeCount - 1
        GetTrusteeList = GetTrusteeList & "!" & trusteeTable(indexTrustee)
    Next
    'WScript.Echo "LIST: " & GetTrusteeList
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get The ACE Mask To check if User/Group has Real Access
'------------------
Function GetACLMask( pMask )
Dim r, sMask
    sMask = ""
    r= pMask AND ADS_RIGHT_GENERIC_ALL
    If r Then
        sMask= "F"
    Else
        r= ( pMask AND ADS_RIGHT_GENERIC_READ ) OR ( pMask AND ADS_RIGHT_READ_CONTROL )
        If r Then
            sMask= sMask & "R"
        End If
         r= ( pMask AND ADS_RIGHT_GENERIC_WRITE ) OR ( pMask AND ADS_RIGHT_WRITE_DAC )
        If r Then
            sMask= sMask & "W"
        End If
        r= pMask AND ADS_RIGHT_GENERIC_EXECUTE
        If r Then
            sMask= sMask & "X"
        End If
        r= pMask AND ADS_RIGHT_DELETE
        If r Then
            sMask= sMask & "D"
        End If
        r= pMask AND ADS_RIGHT_DS_LIST_OBJECT
        If r Then
            sMask= sMask & "L"
        End If
End If
If sMask = "RWDL" Or sMask = "RWXDL" Then
    sMask = "F"
End If
GetACLMask = sMask
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get a User Full Name by is ID (ID=DOMAIN\LOGIN)
'------------------
Function GetUserById ( pUserId )
Dim userDomain , userLogin, userObject
Dim slashPosition, bestTry
On Error Resume Next
    bestTry = GetCacheName( pUserId )
    If bestTry <> "" Then
        GetUserById = bestTry
        Exit Function
    End If
    slashPosition = -1
    slashPosition = InStr(1,pUserId,"\",1)
    If slashPosition > 0 Then
        userDomain = Left(pUserId,slashPosition-1)
        userLogin = Right (pUserId,Len(pUserId)-slashPosition)
        Set userObject = Nothing
        Set userObject = GetObject("WinNT://" & userDomain & "/" & userLogin & ",user")
        If userObject Is Nothing Then
            GetUserById = pUserId
        Else
            GetUserById = userObject.FullName
            Call AddCacheName( pUserId , userObject.FullName )
            Set userObject = Nothing
        End If
    Else
        GetUserById = pUserId
        Call AddCacheName( pUserId , pUserId )
    End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get a Trustee (user or group) Full Name by is ID (ID=DOMAIN\LOGIN or GROUPNAME)
'------------------
Function GetTrusteeById ( pTrusteeId )
Dim userName,groupContent,bestTry
On Error Resume Next
    bestTry = GetCacheName( pTrusteeId )
    If bestTry <> "" Then
        GetTrusteeById = bestTry
        Exit Function
    End If
    userName = GetUserById( pTrusteeId )
    If userName = pTrusteeId Then
        ' Not a user but a group
        groupContent = GetGroupContentById( pTrusteeId )
        GetTrusteeById = groupContent
        Call AddCacheName( pTrusteeId , groupContent )
    Else
        GetTrusteeById = userName
        Call AddCacheName( pTrusteeId , userName )
    End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get a Group Users Full Name by is ID (ID=DOMAIN\GROUPNAME)
'------------------
Function GetGroupContentById ( pGroupId )
Dim groupDomain , groupName, userObject, groupObject, groupContent, userName, bestTry
Dim slashPosition
On Error Resume Next
    bestTry = GetCacheName( pGroupId )
    If bestTry <> "" Then
        pGroupId = bestTry
        Exit Function
    End If
    slashPosition = -1
    slashPosition = InStr(1,pGroupId,"\",1)
    If slashPosition > 0 Then
        groupDomain = Left(pGroupId,slashPosition-1)
        groupName = Right (pGroupId,Len(pUserId)-slashPosition)
        Set groupObject = Nothing
        Set groupObject = GetObject("WinNT://" & groupDomain & "/" & groupName & ",group")
        For Each userObject In groupObject.Members
            userName = GetUserById( groupDomain & "\" & userObject.Name )
            If groupContent = "" Then
                groupContent = userName
            Else
                groupContent = groupContent & Chr(9) & userName
            End If
        Next
        GetGroupContentById = groupContent
        Call AddCacheName( pGroupId , groupContent )
        Set userObject = Nothing
        Set groupObject = Nothing
    Else
        GetGroupContentById = pGroupId
        Call AddCacheName( pGroupId , pGroupId )
    End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Add a Name to the Dictionary Cache (User or Group)
'------------------
Sub AddCacheName( pKey , pValue)
    If Not nameCache.Exists(pKey) Then
        'WScript.Echo "ADD: " & pKey & Chr(9) & pValue
        nameCache.Add pKey,pValue
    End If
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Retrieve a Name to the Dictionary Cache (User or Group)
'------------------
Function GetCacheName( pKey )
    GetCacheName = ""
    If nameCache.Exists(pKey) Then
        GetCacheName = nameCache.Item(pKey)
    End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get Current Domain Name
'------------------
Function GetDefaultDomain()
Dim oNetwork, domainName
On Error Resume Next
    Set oNetwork = WScript.CreateObject("WScript.Network")
    domainName = ""
    domainName = oNetwork.UserDomain
    If domainName = "" Then
        domainName = oNetwork.ComputerName
    End If
    Set oNetwork = Nothing
    GetDefaultDomain = domainName
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get The Server related to the Scan Path
'------------------
Function GetScanServer( pDrive )
Dim oNetwork, serverName, oDrives,I
Dim slashPosition
On Error Resume Next
    Set oNetwork = WScript.CreateObject("WScript.Network")
    serverName = ""
    Set oDrives = oNetwork.EnumNetworkDrives()
For I = 0 To oDrives.Length STEP 2
        WScript.Echo "Drive " & oDrives.Item(I) & " = " & oDrives.Item(I + 1)
        If oDrives.Item(I) = pDrive Then
            slashPosition = InStr(3,oDrives.Item(I+1),"\",1)
            serverName = Left( oDrives.Item(I+1), slashPosition - 1 )
            serverName = Right(serverName,Len(serverName)-2)
            Exit For
        End If
    Next
    Set oNetwork = Nothing
    GetScanServer = serverName
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Get The Server related to the Scan Path
'------------------
Function GetScanShare( pDrive )
Dim oNetwork, shareName, oDrives,I
Dim slashPosition
On Error Resume Next
    Set oNetwork = WScript.CreateObject("WScript.Network")
    serverName = ""
    Set oDrives = oNetwork.EnumNetworkDrives()
For I = 0 To oDrives.Length STEP 2
        WScript.Echo "Drive " & oDrives.Item(I) & " = " & oDrives.Item(I + 1)
        If oDrives.Item(I) = pDrive Then
            shareName = oDrives.Item(I+1)
            Exit For
        End If
    Next
    Set oNetwork = Nothing
    GetScanShare = shareName
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Load the Results into an Excel Spreadsheet
'------------------
Sub ProduceExcelResults()
Dim oInput, iRow, iColumn, S, tFields
Dim oXl
    On Error Resume Next
    oFs.DeleteFile ("C:\TEMP\DIRAGING.XLS")
    Set oXl = WScript.GetObject("Excel.Application")
    If oXl Is Nothing Then
        Set oXl = WScript.CreateObject("Excel.Application")
    End If
    If oXl Is Nothing Then
        oDebug.WriteLine "CANNOT LAUNCH EXCEL"
        Exit Sub
    End If
    oXl.Visible = True
    oXl.Workbooks.Add
    oXl.ActiveSheet.Name="Directories"
    oOutput.Close
    Set oInput = oFs.OpenTextFile(logPath & "DIRAGING.TXT",ForReading,True)
    iRow = 1
    Do While Not oInput.AtEndOfStream
        S = oInput.ReadLine
        tFields = Split(S,Chr(9),-1,1)
        For iColumn = 1 To UBound(tFields) + 1
            oXl.ActiveSheet.Cells(iRow,iColumn).Value = tFields(iColumn-1)
        Next
        iRow = iRow + 1
    Loop
    oInput.Close
    oXl.ActiveWorkbook.SaveAs ("C:\TEMP\DIRAGING.XLS")
    oXl.ActiveWorkbook.Close
    oXL.Application.quit
End Sub
Function BuildArchivePath ( pPath )
Dim d1, p
    d1= Right(pPath,Len(pPath)-2)
    d1 = Left(d1,Len(d1)-1)
    d1 = archivePath & d1
    d1 = NoLastSlash(d1)
    p = InStrRev(d1,"\",-1,1) - 1
    BuildArchivePath = Left(d1,p)
    
End Function
Sub MkdirArchivePath ( pPath )
Dim d1 , dTable, i, dName
    d1= Right(pPath,Len(pPath)-2)
    d1 = Left(d1,Len(d1)-1)
    d1 = archivePath & d1
    dTable = Split(d1,"\",-1,1)
    dName = dTable(0)
    For i=1 To UBound(dTable)
        dName = dName & "\" & dTable(i)
    '    oArchive.WriteLine "MKDIR " & Chr(34) & dName & Chr(34)
    Next
    oArchive.WriteLine "MKDIR " & Chr(34) & dName & Chr(34)
End Sub
Function NoLastSlash(pPath)
    If Right(pPath,1) = "\" Then
        NoLastSlash = Left(pPath,Len(pPath)-1)
    Else
        NoLastSlash = pPath
    End If
End Function
Sub PushDirectory( pPath )
    directoryStack( stackDepth ) = pPath
    oDebug.WriteLine "PUSH: " & directoryStack( stackDepth )
    stackDepth = stackDepth + 1
End Sub
Function PopDirectory()
    If stackDepth > 0 Then
        PopDirectory = directoryStack( stackDepth - 1 )
        stackDepth = stackDepth - 1
        oDebug.WriteLine "POP: " & directoryStack( stackDepth )
    Else
        PopDirectory = ""
    End If
End Function
Function IsEmptyStack()
    If stackDepth > 0 Then
        IsEmptyStack = False
    Else
        IsEmptyStack = True
    End If
End Function
