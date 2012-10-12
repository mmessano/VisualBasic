'===============================================================================
'- generate scripts to rebuild all objects in [SQL server name]/[SQL database name]
'- check scripts into SourceSafe whenever current script is different than VSS's
'
'Revisions: When, Who, Why
' 03/20/2001    Rich Cowley (cowley_rich at hotmail.com)   Original Version
'
' V.1
' Extensivley Modified to be a VB Script (was a VB BAS module)
' David Jackson
'
' V.2
' Changed to takes parameters instead of hardcoded server names &tc.
' David Jackson
'
' V2.1
' Fixed bug where was passing in VSS_ROOTPROJECT_NAME instead of SCRIPT_DIRECTORY to 
' MakeSureDirectoryTreeExists function
' David Jackson - 14:24 21/10/2004
'
' Instructions.
' To use this script, you need to create a project in VSS & create various sub projects.
' 
'       -/$
'         |-Databases
'               |-pubs
'                   |-Defaults
'                   |-Rules
'                   |-StoredProcedures 
'                   |-Tables 
'                   |-UserDefinedDataTypes 
'                   |-UserDefinedFunctions 
'                   |-Views
'
'
' You then need to create a batch file to pass in the required parameters.
'
'	@Echo Off
'	::SQLServer1
'	cscript dbtovss2.vbs -w C:\Temp\pubs -r $/Databases/pubs/ -i \\VSSServer\VSS\Data\srcsafe.ini -u Admin -p secret -s SQLServer1 -d pubs
'	::call it again for a different SQL db
'	cscript dbtovss2.vbs -w C:\Temp\dbname2 -r $/Databases/dbname2/ -i \\VSSServer\VSS\Data\srcsafe.ini -u Admin -p secret -s SQLServer1 -d dbname2
'	::End of Batch file
'
' This  needs to run on a PC with both a VSS & SQL client installed, as well as having TLBINF32.DLL registered. 
' To register TLBINF32.DLL, open a command line in the folder you have copied TLBINF32.DLL to and type:
'
' Regsvr32 TLBINF32.DLL
'

' TLBINF32.DLL is shipped on the Visual Studio 6.0 and Visual Basic 6.0 CDs.  
' According to Microsoft it is redistributable,
' so look for it at http://glossopian.co.uk/Uploads/Main/TLBINF32.zip.
'
' If you create the structure above, this script will script off the pubs db 'out of the box'
'
'============================================================================================
Option Explicit

Dim SCRIPT_DIRECTORY, VSS_ROOTPROJECT_NAME, VSS_INI_PATH, VSS_USERNAME, VSS_PASSWORD, SQL_SERVERNAME, SQL_DBNAME
'General Constants
SCRIPT_DIRECTORY       = GetArgs( "w", "C:\Temp" ) 'working directory on local machine

'SourceSafe Constants
VSS_ROOTPROJECT_NAME   = GetArgs( "r", "$/Product Operations/Databases/dbamaint/" )
VSS_INI_PATH           = GetArgs( "i", "C:\Program Files\Microsoft Visual Studio\Common\VSS\srcsafe.ini" )
VSS_USERNAME           = GetArgs( "u", "Admin" )
VSS_PASSWORD           = GetArgs( "p", "Secret" )

'SQL DMO Constants
SQL_SERVERNAME         = GetArgs( "s", "(local)" )
SQL_DBNAME             = GetArgs( "d", "pubs" )

Const SQLDMOScript_Default = 4
Const SQLDMOScript_Drops = 1
Const SQLDMOScript_Triggers = 16
Const SQLDMOScript_Indexes = 73736

'FileSystemObject constants
Const ForReading = 1
Const ForWriting = 2

MakeSureDirectoryTreeExists(SCRIPT_DIRECTORY)

Dim oSQLServer      'As SQLDMO.SQLServer2
Dim oDatabase       'As SQLDMO.Database2
Dim oDatabaseObject
Dim sScript         'As String
Dim oFSO            'As Scripting.FileSystemObject
Dim oFolder         'As Folder
Dim oFile           'As File
Dim oTS             'As Scripting.TextStream
Dim sCurrDirectory  'As String
Dim sFileName       'As String
Dim i               'As Integer
Dim iScriptOptions  'As Integer
Dim sObjectType     'As String

Dim oVSSDatabase    'As SourceSafeTypeLib.VSSDatabase
Dim oVSSItem        'As VSSItem
Dim sVSSItemPath    'As String
Dim sVSSLabel       'As String 'label used when checking items back in
Dim sCheckedScript  'As String

Dim sTempScript1    'As String
Dim sTempScript2    'As String
Dim itemCounter     'As Integer

'============================================================================================
'Script every object in the database (tables, views, SPs, etc.)
'If current script is different than SourceSafe's, check in new version.
'Add new SourceSafe item if current script was not in SourceSafe to begin with.
'============================================================================================

    'set up SourceSafe environment
    Set oVSSDatabase = ImportObject("SourceSafe")
    oVSSDatabase.Open VSS_INI_PATH, VSS_USERNAME, VSS_PASSWORD
    'Set counter up
    itemCounter = 0
    'establish SQL Server and FSO environments
    Set oSQLServer = ImportObject("SQLDMO.SQLServer2")
    oSQLServer.LoginSecure = True
    oSQLServer.Connect SQL_SERVERNAME

    Set oDatabase = oSQLServer.Databases(SQL_DBNAME)
    Set oFSO = wScript.CreateObject("Scripting.FileSystemObject")

    '----------------
    'Get all objects
    '----------------
    GetObjects("Tables")
    GetObjects("Views")
    GetObjects("StoredProcedures")
    GetObjects("Rules")
    GetObjects("Defaults")
    GetObjects("UserDefinedDatatypes")
    GetObjects("UserDefinedFunctions")

    '--------
    'clean up
    '--------
    oSQLServer.Close
    Set oSQLServer = Nothing
    Set oDatabase = Nothing

    Set oVSSItem = Nothing
    Set oVSSDatabase = Nothing

    'toss the working folders
    oFSO.DeleteFolder SCRIPT_DIRECTORY, True

    Set oFSO = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
    Set oTS = Nothing
    On Error Goto 0
    'process complete!
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")

    i = WshShell.Popup (itemCounter & " SQL Objects successfully rolled to SourceSafe.", 10, "VSS autorun DB Objects", 64)
    Set WshShell = Nothing

'----------------------------------------------------------------------
Sub GetObjects(byVal objType)
        '----------------------------------------------------------------------
            sObjectType = objType
            sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/"
            Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
            'set up clean working directory
            sCurrDirectory = SCRIPT_DIRECTORY & "\" & sObjectType
            'On Error Resume Next
            If oFSO.FolderExists(sCurrDirectory) Then
                oFSO.DeleteFolder sCurrDirectory, True
            End If
            Set oFolder = oFSO.CreateFolder(sCurrDirectory)

            'cycle through the objects
            Select Case objType
                    Case "Tables"
                        For Each oDatabaseObject In oDatabase.Tables
                            If oDatabaseObject.SystemObject Then
                                'do nothing (bypass system objects)
                            Else
                                iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops + SQLDMOScript_Triggers + SQLDMOScript_Indexes
                                sScript = oDatabaseObject.Script(iScriptOptions)
                                sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                                sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                                On Error Resume Next
                                Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                                If Err = 0 Then 'item is already on SourceSafe
                                    oVSSItem.Checkout "Checked out by automated process", sFileName
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.WriteLine (sScript)
                                    oTS.Close
                                    oVSSItem.Checkin "Schema Altered", sFileName
                                    itemCounter = itemCounter + 1
                                Else 'item does not yet exist on SourceSafe; add it
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.Write (sScript)
                                    oTS.Close
                                    Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                    oVSSItem.Add sFileName
                                    itemCounter = itemCounter + 1
                                End If
                                Set oVSSItem = Nothing
                            End If
                        Next

                    Case "Views"
                        For Each oDatabaseObject In oDatabase.Views
                            If oDatabaseObject.SystemObject Then
                                'do nothing (bypass system objects)
                            Else
                                iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops + SQLDMOScript_Triggers + SQLDMOScript_Indexes
                                sScript = oDatabaseObject.Script(iScriptOptions)
                                sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                                sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                                On Error Resume Next
                                Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                                If Err = 0 Then 'item is already on SourceSafe
                                    oVSSItem.Checkout "Checked out by automated process", sFileName
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.WriteLine (sScript)
                                    oTS.Close
                                    oVSSItem.Checkin "Schema Altered", sFileName
                                    itemCounter = itemCounter + 1
                                Else 'item does not yet exist on SourceSafe; add it
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.Write (sScript)
                                    oTS.Close
                                    Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                    oVSSItem.Add sFileName
                                    itemCounter = itemCounter + 1
                                End If
                                Set oVSSItem = Nothing
                            End If
                        Next

                    Case "StoredProcedures"
                        For Each oDatabaseObject In oDatabase.StoredProcedures
                            If oDatabaseObject.SystemObject Then
                                'do nothing (bypass system objects)
                            Else
                                iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops
                                sScript = oDatabaseObject.Script(iScriptOptions)
                                sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                                sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                                On Error Resume Next
                                Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                                If Err = 0 Then 'item is already on SourceSafe
                                    oVSSItem.Checkout "Checked out by automated process", sFileName
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.WriteLine (sScript)
                                    oTS.Close
                                    oVSSItem.Checkin "Schema Altered", sFileName
                                    itemCounter = itemCounter + 1
                                Else 'item does not yet exist on SourceSafe; add it
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.Write (sScript)
                                    oTS.Close
                                    Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                    oVSSItem.Add sFileName
                                    itemCounter = itemCounter + 1
                                End If
                                Set oVSSItem = Nothing
                            End If
                        Next

                    Case "Defaults"
                        For Each oDatabaseObject In oDatabase.Defaults
                                iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops
                                sScript = oDatabaseObject.Script(iScriptOptions)
                                sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                                sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                                On Error Resume Next
                                Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                                If Err = 0 Then 'item is already on SourceSafe
                                    oVSSItem.Checkout "Checked out by automated process", sFileName
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.WriteLine (sScript)
                                    oTS.Close
                                    oVSSItem.Checkin "Schema Altered", sFileName
                                    itemCounter = itemCounter + 1
                                Else 'item does not yet exist on SourceSafe; add it
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.Write (sScript)
                                    oTS.Close
                                    Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                    oVSSItem.Add sFileName
                                    itemCounter = itemCounter + 1
                                End If
                                Set oVSSItem = Nothing
                        Next

                    Case "Rules"
                        For Each oDatabaseObject In oDatabase.Rules
                            If oDatabaseObject.SystemObject Then
                                'do nothing (bypass system objects)
                            Else
                                iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops
                                sScript = oDatabaseObject.Script(iScriptOptions)
                                sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                                sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                                On Error Resume Next
                                Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                                If Err = 0 Then 'item is already on SourceSafe
                                    oVSSItem.Checkout "Checked out by automated process", sFileName
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.WriteLine (sScript)
                                    oTS.Close
                                    oVSSItem.Checkin "Schema Altered", sFileName
                                    itemCounter = itemCounter + 1
                                Else 'item does not yet exist on SourceSafe; add it
                                    Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                    oTS.Write (sScript)
                                    oTS.Close
                                    Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                    oVSSItem.Add sFileName
                                    itemCounter = itemCounter + 1
                                End If
                                Set oVSSItem = Nothing
                            End If
                        Next

                    Case "UserDefinedDataTypes"
                        For Each oDatabaseObject In oDatabase.UserDefinedDataTypes
                            iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops
                            sScript = oDatabaseObject.Script(iScriptOptions)
                            sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                            sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                            On Error Resume Next
                            Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                            If Err = 0 Then 'item is already on SourceSafe
                                oVSSItem.Checkout "Checked out by automated process", sFileName
                                Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                oTS.WriteLine (sScript)
                                oTS.Close
                                oVSSItem.Checkin "Schema Altered", sFileName
                                itemCounter = itemCounter + 1
                            Else 'item does not yet exist on SourceSafe; add it
                                Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                oTS.Write (sScript)
                                oTS.Close
                                Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                oVSSItem.Add sFileName
                                itemCounter = itemCounter + 1
                            End If
                            Set oVSSItem = Nothing
                        Next

                    Case "UserDefinedFunctions"
                        For Each oDatabaseObject In oDatabase.UserDefinedFunctions
                            iScriptOptions = SQLDMOScript_Default + SQLDMOScript_Drops
                            sScript = oDatabaseObject.Script(iScriptOptions)
                            sVSSItemPath = VSS_ROOTPROJECT_NAME & sObjectType & "/" & oDatabaseObject.Name & ".sql"
                            sFileName = sCurrDirectory & "\" & oDatabaseObject.Name & ".sql"
                            On Error Resume Next
                            Set oVSSItem = oVSSDatabase.VSSItem(sVSSItemPath)
                            If Err = 0 Then 'item is already on SourceSafe
                                oVSSItem.Checkout "Checked out by automated process", sFileName
                                Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                oTS.WriteLine (sScript)
                                oTS.Close
                                oVSSItem.Checkin "Schema Altered", sFileName
                                itemCounter = itemCounter + 1
                            Else 'item does not yet exist on SourceSafe; add it
                                Set oTS = oFSO.OpenTextFile(sFileName, ForWriting, True)
                                oTS.Write (sScript)
                                oTS.Close
                                Set oVSSItem = oVSSDatabase.VSSItem(VSS_ROOTPROJECT_NAME & sObjectType & "/")
                                oVSSItem.Add sFileName
                                itemCounter = itemCounter + 1
                            End If
                            Set oVSSItem = Nothing
                        Next

                End Select
End Sub

'----------------------------------------------------------------------
Function ImportObject(sClass)
'----------------------------------------------------------------------
' PURPOSE: Given a classname, this function will:
' + return a reference to the object
' + Import Typelib constants into global namespace
' DEPENDENCY: TLBINF32.DLL must be present and registered
' Derived from Michael Harris' Typelib extraction HTA
' WARNING: Some TLBs contain hundreds of constants!
' Alex K. Angelopoulos posted this on 10/31/02.
'
'Sample usage:
'<CODE>
'Set objIE = ImportObject("InternetExplorer.Application")
'</CODE>
'which is identical to the following code lines, except that
' all approximately 80 constants are defined.
'<CODE>
'Set objIE = CreateObject("InternetExplorer.Application")
'CONST CSC_UPDATECOMMANDS = -1
' ... many more CONST statements
'CONST SWFO_COOKIEPASSED = 4
'</CODE>
'
' Edited and modified to allow repeated calls with
'   the same class.   Paul Randall 11/21/02

Dim objTLIA   'TypeLib Info Application; TLBinf32.dll
Dim objTLII   'TypeLib Interface Info object for the parent of
      ' the object created from the specified class.
      ' Contains a collection of enumeration objects
Dim objCEnum  'One of the enumeration objects
      ' Contains a collection of constant objects
Dim objConstant 'One constant object in the enumeration object
Dim objObject  'The object specified by the class string passed
      ' to this routine.

'strMsg for obtaining the list of constants and their values
Dim strMsg   'List of constants and their values
strMsg = "Typelib constants for: " & sClass & vbcrlf & vbcrlf

Set objObject = CreateObject(sClass)
Set objTLIA = CreateObject("TLI.TLIApplication")
Set objTLII = objTLIA.InterfaceInfoFromObject(objObject).Parent
For Each objCEnum in objTLII.Constants
 ' We only want them if they are visible
 If Left(objCEnum.Name, 1)<>"_" Then
  strMsg = strMsg & "EnumName: " & objCEnum.Name & _
   " contains " & objCEnum.Members.count & " items." & vbcrlf
  For Each objConstant In objCEnum.Members
   strMsg = strMsg & objConstant.name & " = " & objConstant.value & vbcrlf
   On Error Resume Next
   ExecuteGlobal "CONST " & objConstant.Name & "=" & objConstant.Value
   if Err.Number = 1041 then
    if eval(objConstant.Name & "=" & objConstant.Value) then
     'Ignore unchanged values
    else
     MsgBox "Unexpected new value for TypeLib constant" & vbcrlf & _
      "in Function ImportObject(" & sClass & ")" & vbcrlf & _
      "Constant name = " & objConstant.Name & vbcrlf & _
      "Old value = " & eval(objConstant.Name) & vbcrlf & _
      "New value = " & objConstant.Value & vbcrlf & _
      vbcrlf & "Quitting"
     WScript.Quit
    end if
   elseif Err.Number <> 0 then
    MsgBox "Unexpected error in " & _
     "Function ImportObject(" & sClass & ")" & vbcrlf & _
     "Error Number = " & Err.Number & vbcrlf & _
     "Error Description = " & Err.Description & vbcrlf & _
     "Error Source = " & Err.Source & vbcrlf & _
     vbcrlf & "Quitting"
    WScript.Quit
   end if
   On Error GoTo 0
  Next
 else
  strMsg = strMsg & "EnumName: " & objCEnum.Name & _
   " contains " & objCEnum.Members.count & " hidden items." & vbcrlf
 End If
Next
' WriteFile "C:\LogFile.txt", strMsg
Set ImportObject = objObject
End Function

'-----------------------------------------------------
Function GetArgs( sSwitch, sDefaultValue )
'-----------------------------------------------------
' Checks the command line arguments for a given switch and returns the associated
' string, if found. If not found, the defaultValue is returned instead.
dim ArgCount, bMatch
ArgCount = 0
bMatch = 0
do while ArgCount < WScript.arguments.length
    if Eval((WScript.arguments.item(ArgCount)) = ("-" + (sSwitch))) Or Eval((WScript.arguments.item(ArgCount)) = ("/" + (sSwitch))) then
        bMatch = 1
        Exit do
    else
        ArgCount = ArgCount + 1
    end if
Loop
if ( bMatch = 1 ) then
        GetArgs = ( WScript.arguments.item(ArgCount + 1) )
    else
        GetArgs = ( sDefaultValue )
end if

End Function
    
'----------------------------------------------------------------------
Function MakeSureDirectoryTreeExists(dirName)
'----------------------------------------------------------------------
'like it says on the tin
    Dim aFolders, newFolder, i
    Dim oFS
    
    Set oFS = CreateObject("Scripting.FileSystemObject")

   ' Check the folder's existence
   If Not oFS.FolderExists(dirName) Then
      ' Split the various components of the folder's name
      aFolders = split(dirName, "\")

      ' Get the root of the drive
      newFolder = oFS.BuildPath(aFolders(0), "\")

      ' Scan the various folder and create them
      For i = 1 To UBound(aFolders)
         newFolder = oFS.BuildPath(newFolder, aFolders(i))

         If Not oFS.FolderExists(newFolder) Then
            oFS.CreateFolder newFolder
         End If
      Next
   End If
       Set oFS = Nothing
End Function







