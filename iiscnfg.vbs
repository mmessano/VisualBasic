'
' Copyright (c) Microsoft Corporation.  All rights reserved.
'
' VBScript Source File 
'
' Script Name: IIsCnfg.vbs
'

Option Explicit
On Error Resume Next

' Error codes
Const ERR_OK              = 0
Const ERR_GENERAL_FAILURE = 1

'''''''''''''''''''''
' Messages
Const L_ConfImported_Text       = "Configuration imported from %1 in file"
Const L_ConfImportedp2_Text     = "%1 to %2 in the Metabase."
Const L_ConfExported_Text       = "Configuration exported from %1 to file %2."
Const L_MDSaved_Text            = "Metadata successfully flushed to disk."

Const L_Error_ErrorMessage                   = "Error &H%1: %2"
Const L_GetComputerObject_ErrorMessage       = "Could not get computer object"
Const L_Import_ErrorMessage                  = "Error while importing configuration."
Const L_Export_ErrorMessage                  = "Error while exporting configuration."
Const L_SaveData_ErrorMessage                = "Error while flushing metabase."
Const L_OnlyOneOper_ErrorMessage             = "Please specify only one operation at a time."
Const L_ScriptHelper_ErrorMessage            = "Could not create an instance of the"
Const L_ScriptHelperp2_ErrorMessage          = "IIsScriptHelper object."
Const L_ChkScpHelperReg_ErrorMessage         = "Please register the Microsoft.IIsScriptHelper"
Const L_ChkScpHelperRegp2_ErrorMessage       = "component."
Const L_CmdLib_ErrorMessage                  = "Could not create an instance of the CmdLib object."
Const L_ChkCmdLibReg_ErrorMessage            = "Please register the Microsoft.CmdLib component."
Const L_WMIConnect_ErrorMessage              = "Could not connect to WMI provider."
Const L_RequiredArgsMissing_ErrorMessage     = "Required arguments are missing."
Const L_FileExpected_ErrorMessage            = "Argument is a folder path while expecting a file"
Const L_FileExpectedp2_ErrorMessage          = "path."
Const L_ParentFolderDoesntExist_ErrorMessage = "Parent folder doesn't exist."
Const L_FileDoesntExist_ErrorMessage         = "Input file doesn't exist."
Const L_FileAlreadyExist_ErrorMessage        = "Export file specified already exists."
Const L_NotEnoughParams_ErrorMessage         = "Not enough parameters."
Const L_InvalidSwitch_ErrorMessage           = "Invalid switch: %1"
Const L_IncorrectPassword_ErrorMessage       = "The password specified is incorrect."
Const L_InvalidXML_ErrorMessage              = "The import file appears to contain invalid XML."
Const L_Admin_ErrorMessage                   = "You cannot run this command because you are not an"
Const L_Adminp2_ErrorMessage                 = "administrator on the server you are trying to configure."
Const L_DriveLetter_Message                  = "Mapping local drive %1 to admin share on server %2"
Const L_Shell_ErrorMessage                   = "Could not create an instance of the WScript.Shell"
Const L_Shellp2_ErrorMessage                 = "object."
Const L_FS_ErrorMessage                      = "Could not create an instance of the"
Const L_FSp2_ErrorMessage                    = "Scripting.FileSystemObject object."
Const L_Network_ErrorMessage                 = "Could not create an instance of the"
Const L_Networkp2_ErrorMessage               = "WScript.Network object."
Const L_BackingUp_Message                    = "Backing up server %1"
Const L_Restoring_Message                    = "Restoring on server %1"
Const L_Backup_ErrorMessage                  = "Failure creating backup."
Const L_BackupComplete_Message               = "Backup complete."
Const L_Restore_ErrorMessage                 = "Failure restoring backup."
Const L_RestoreComplete_Message              = "Restore complete."
Const L_UnMap_Message                        = "Unmapping local drive %1"
Const L_NoDrive_ErrorMessage                 = "No drives available for mapping on local machine."
Const L_Copy_Message                         = "Copying backup files..."
Const L_Copy_ErrorMessage                    = "Error copying files."
Const L_ReturnVal_ErrorMessage               = "Call returned with code %1"
Const L_CopyComplete_Message                 = "Copy operation complete."
Const L_AuditAlreadyEnabled_Message          = "Auditing is already enabled on %1. No action performed."
Const L_AuditAlreadyDisabled_Message         = "Auditing is already disabled on %1. No action performed."
Const L_EnableAuditFail_ErrorMessage         = "Failed to enable auditing on %1."
Const L_DisableAuditFail_ErrorMessage        = "Failed to disable auditing on %1."
Const L_AuditEnabled_Message                 = "Auditing was enabled on %1."
Const L_AuditDisabled_Message                = "Auditing was disabled on %1."
Const L_GetObject_ErrorMessage               = "Error retrieving ADSI object %1."
Const L_GetDataPaths_ErrorMessage            = "Error looking for the AdminACL property in the metabase."

'''''''''''''''''''''
' Help
Const L_Empty_Text     = ""

' General help messages
Const L_SeeHelp_Message             = "Type IIsCnfg /? for help."
Const L_SeeImportHelp_Message       = "Type IIsCnfg /import /? for help."
Const L_SeeExportHelp_Message       = "Type IIsCnfg /export /? for help."
Const L_SeeEnableAuditHelp_Message  = "Type IIsCnfg /EnableAudit /? for help."
Const L_SeeDisableAuditHelp_Message = "Type IIsCnfg /DisableAudit /? for help."

Const L_Help_HELP_General01_Text   = "Description: Import and export IIS configuration."
Const L_Help_HELP_General02_Text   = "Syntax: IIsCnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_General03_Text   = "        /<operation> [arguments]"
Const L_Help_HELP_General04_Text   = "Parameters:"
Const L_Help_HELP_General05_Text   = ""
Const L_Help_HELP_General06_Text   = "Value                   Description"
Const L_Help_HELP_General07_Text   = "/s <server>             Connect to machine <server>."
Const L_Help_HELP_General07p2_Text = "                        [Default: this system]"
Const L_Help_HELP_General08_Text   = "/u <username>           Connect as <domain>\<username> or"
Const L_Help_HELP_General09_Text   = "                        <username>. [Default: current user]"
Const L_Help_HELP_General10_Text   = "/p <password>           Password for the <username> user."
Const L_Help_HELP_General11_Text   = "<operation>             /import       Import configuration from"
Const L_Help_HELP_General11p1_Text = "                                      a configuration file."
Const L_Help_HELP_General12_Text   = "                        /export       Export configuration into"
Const L_Help_HELP_General12p1_Text = "                                      a configuration file."
Const L_Help_HELP_General13_Text   = "                        /copy         Copy configuration from"
Const L_Help_HELP_General13p1_Text = "                                      one machine to another."
Const L_Help_HELP_General14_Text   = "                        /EnableAudit  Enable auditing on a"
Const L_Help_HELP_General14p1_Text = "                                      certain metabase path."
Const L_Help_HELP_General15_Text   = "                        /DisableAudit Disable auditing on a"
Const L_Help_HELP_General15p1_Text = "                                      certain metabase path."
Const L_Help_HELP_General22_Text   = "For detailed usage:"
Const L_Help_HELP_General23_Text   = "IIsCnfg /import /?"
Const L_Help_HELP_General24_Text   = "IIsCnfg /export /?"
Const L_Help_HELP_General25_Text   = "IIsCnfg /copy /?"
Const L_Help_HELP_General26_Text   = "IIsCnfg /save /?"
Const L_Help_HELP_General27_Text   = "IIsCnfg /EnableAudit /?"
Const L_Help_HELP_General28_Text   = "IIsCnfg /DisableAudit /?"

' Common help messages
Const L_Help_HELP_Common13_Text   = "/d <DecryptPass>        Specifies the password used to"
Const L_Help_HELP_Common13p1_Text = "                        decrypt encrypted configuration data."
Const L_Help_HELP_Common13p2_Text = "                        [Default: """"]"   
Const L_Help_HELP_Common14_Text   = "/f <File>               Configuration file."
Const L_Help_HELP_Common15_Text   = "/sp <SourcePath>        The full metabase path to start"
Const L_Help_HELP_Common15p1_Text = "                        reading from the configuration file."
Const L_Help_HELP_Common21_Text   = "Examples:"

' Copy help messages
Const L_Help_HELP_Copy1_Text      = "Description:  Copy configuration from a source server to a"
Const L_Help_HELP_Copy1p2_Text    = "              target server."
Const L_Help_HELP_Copy2_Text      = "Syntax: iiscnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_Copy2p2_Text    = "        /copy  /ts <target server> /tu <target user>"
Const L_Help_HELP_Copy2p3_Text    = "        /tp <target password>"
Const L_Help_HELP_Copy3_Text      = "Parameters:"
Const L_Help_HELP_Copy4_Text      = "Value		           Description"
Const L_Help_HELP_Copy5_Text      = "/s <server>           Connect to machine <server>"
Const L_Help_HELP_Copy5p2_Text    = "                      [Default: this system]"
Const L_Help_HELP_Copy6_Text      = "/u <username>         Connect as <domain>/<username>"
Const L_Help_HELP_Copy7_Text      = "                      or <username> [Default: current user]"
Const L_Help_HELP_Copy8_Text      = "/p <password>         Password for the <username> user"
Const L_Help_HELP_Copy9_Text      = "/ts                   Target server to copy configuration to"
Const L_Help_HELP_Copy10_Text     = "/tu                   Username to use when connecting to the"
Const L_Help_HELP_Copy10p2_Text   = "                      target server"
Const L_Help_HELP_Copy11_Text     = "/tp                   Password to use when connecting to the"
Const L_Help_HELP_Copy11p2_Text   = "                      target server"
Const L_Help_HELP_Copy12_Text     = "Examples:"
Const L_Help_HELP_Copy13_Text     = "IIsCnfg /copy /ts TargetServer /tu Administrator /tp Pk$^("
Const L_Help_HELP_Copy14_Text     = "IIsCnfg /s SourceServer /u Administrator /p Kj30W /copy"
Const L_Help_HELP_Copy14p2_Text   = "        /ts TargetServer /tu Administrator /tp Pk$^j"

' Import help messages
Const L_Help_HELP_Import01_Text   = "Description: Import configuration from a configuration file."
Const L_Help_HELP_Import02_Text   = "Syntax: IIsCnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_Import02p1_Text = "        /import [/d <DeCryptPass>] /f <File> /sp <SourcePath>"
Const L_Help_HELP_Import02p2_Text = "        /dp <DestPath> [/inherited] [/children] [/merge]"
Const L_Help_HELP_Import16_Text   = "/dp <DestPath>          The metabase path destination for"
Const L_Help_HELP_Import16p1_Text = "                        imported properties.  If the keytype"
Const L_Help_HELP_Import16p2_Text = "                        of the SourcePath and the DestPath do"
Const L_Help_HELP_Import16p3_Text = "                        not match, an error occurs."
Const L_Help_HELP_Import17_Text   = "/inherited              Import inherited settings if set."
Const L_Help_HELP_Import18_Text   = "/children               Import configuration for child nodes."
Const L_Help_HELP_Import19_Text   = "/merge                  Merge imported configuration with"
Const L_Help_HELP_Import19p1_Text = "                        existing configuration."
Const L_Help_HELP_Import22_Text   = "IIsCnfg /import /f c:\config.xml /sp /lm/w3svc/5/Root/401Kapp"
Const L_Help_HELP_Import22p1_Text = "        /dp /lm/w3svc/1/Root/401Kapp"

' Export help messages
Const L_Help_HELP_Export01_Text   = "Description: Export configuration into a configuration file."
Const L_Help_HELP_Export02_Text   = "Syntax: IIsCnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_Export02p1_Text = "        /export [/d <DeCryptPass>] /f <File> /sp <SourcePath>"
Const L_Help_HELP_Export02p2_Text = "         [/inherited] [/children]"
Const L_Help_HELP_Export17_Text   = "/inherited              Export inherited settings if set."
Const L_Help_HELP_Export18_Text   = "/children               Export configuration for child nodes."
Const L_Help_HELP_Export22_Text   = "IIsCnfg /export /f c:\config.xml /sp /lm/w3svc/5/Root/401Kapp"

' Save help messages
Const L_Help_HELP_Save01_Text   = "Description:  Save configuration to disk."
Const L_Help_HELP_Save02_Text   = "Syntax: IIsCnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_Save02p2_Text = "        /save"
Const L_Help_HELP_Save22_Text   = "IIsCnfg /save"
Const L_Help_HELP_Save23_Text   = "IIsCnfg /s SourceServer /u Administrator /p Kj30W /save"

' EnableAudit help messages
Const L_Help_HELP_EnableAudit01_Text   = "Description:  Enable auditing on the metabase path indicated."
Const L_Help_HELP_EnableAudit02_Text   = "Syntax: IIsCnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_EnableAudit02p2_Text = "        /EnableAudit <path> [/r]"
Const L_Help_HELP_EnableAudit16_Text   = "<path>                  The full metabase path to enable"
Const L_Help_HELP_EnableAudit16p1_Text = "                        auditing on."
Const L_Help_HELP_EnableAudit17_Text   = "/r                      Recursively enable auditing starting"
Const L_Help_HELP_EnableAudit17p1_Text = "                        from <path>."
Const L_Help_HELP_EnableAudit22_Text   = "IIsCnfg /EnableAudit /w3svc/1/root"
Const L_Help_HELP_EnableAudit23_Text   = "IIsCnfg /EnableAudit / /r"
Const L_Help_HELP_EnableAudit24_Text   = "IIsCnfg /s SourceServer /u Administrator /p Kj30W /EnableAudit /w3svc/2 /r"

' DisableAudit help messages
Const L_Help_HELP_DisableAudit01_Text   = "Description:  Disable auditing on the metabase path indicated."
Const L_Help_HELP_DisableAudit02_Text   = "Syntax: IIsCnfg [/s <server> [/u <username> [/p <password>]]]"
Const L_Help_HELP_DisableAudit02p2_Text = "        /DisableAudit <path> [/r]"
Const L_Help_HELP_DisableAudit16_Text   = "<path>                  The full metabase path to disable"
Const L_Help_HELP_DisableAudit16p1_Text = "                        auditing on."
Const L_Help_HELP_DisableAudit17_Text   = "/r                      Recursively disable auditing starting"
Const L_Help_HELP_DisableAudit17p1_Text = "                        from <path>."
Const L_Help_HELP_DisableAudit22_Text   = "IIsCnfg /DisableAudit /w3svc/1/root"
Const L_Help_HELP_DisableAudit23_Text   = "IIsCnfg /DisableAudit / /r"
Const L_Help_HELP_DisableAudit24_Text   = "IIsCnfg /s SourceServer /u Administrator /p Kj30W /DisableAudit /w3svc/2 /r"

''''''''''''''''''''''''
' Operation codes
Const OPER_IMPORT        = 1
Const OPER_EXPORT        = 2
Const OPER_COPY          = 3
Const OPER_SAVE          = 4
Const OPER_ENABLE_AUDIT  = 5
Const OPER_DISABLE_AUDIT = 6

' Import/Export flags
Const IMPORT_EXPORT_INHERITED = 1
Const IMPORT_EXPORT_NODE_ONLY = 2
Const IMPORT_EXPORT_MERGE     = 4

' Used in the call to GetDataPaths
Const IIS_DATA_INHERIT = 1

Const ADMINISTRATORS_GROUP_SID = "S-1-5-32-544"

'
' Main block
'
Dim oScriptHelper, oCmdLib
Dim strServer, strUser, strPassword, strSite
Dim strTarServer, strTarUser, strTarPassword
Dim strFile, strDecPass, strSourcePath, strDestPath, strMBPath
Dim intOperation, intResult, intFlags
Dim bRecurseAuditOperation
Dim aArgs, arg
Dim strCmdLineOptions
Dim oError

' Default values
strServer = "."
strUser = ""
strPassword = ""
strTarServer = ""
strTarUser = ""
strTarPassword = ""
intOperation = 0
strFile = ""
strDecPass = ""
strSourcePath = ""
strDestPath = ""
strMBPath = ""
intFlags = IMPORT_EXPORT_NODE_ONLY
bRecurseAuditOperation = False

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
    WScript.Echo L_ScriptHelperp2_ErrorMessage
    WScript.Echo L_ChkScpHelperReg_ErrorMessage    
    WScript.Echo L_ChkScpHelperRegp2_ErrorMessage    
    WScript.Quit(ERR_GENERAL_FAILURE)
End If

Set oScriptHelper.ScriptHost = WScript

' Check if we are being run with cscript.exe instead of wscript.exe
oScriptHelper.CheckScriptEngine

' Minimum number of parameters must exist
If WScript.Arguments.Count < 1 Then
    WScript.Echo L_RequiredArgsMissing_ErrorMessage
	WScript.Echo L_SeeHelp_Message
    WScript.Quit(ERR_GENERAL_FAILURE)
End If

strCmdLineOptions = "[server:s:1;user:u:1;password:p:1];decpass:d:1;file:f:1;sourcepath:sp:1;" & _
                    "inherited:i:0;children:c:0;[import::0;destpath:dp:1;merge:m:0];save::0;" & _
                    "export::0;[copy::0;targetserver:ts:1;targetuser:tu:1;targetpassword:tp:1];" & _
                    "[enableaudit:ea:1;recurse:r:0];[disableaudit:da:1;recurse:r:0]"
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
        Case "server"
            ' Server information
            strServer = oScriptHelper.GetSwitch(arg)

        Case "user"
            ' User information
            strUser = oScriptHelper.GetSwitch(arg)

        Case "password"
            ' Password information
            strPassword = oScriptHelper.GetSwitch(arg)
            
        Case "targetserver"
            ' Server information
            strTarServer = oScriptHelper.GetSwitch(arg)

        Case "targetuser"
            ' User information
            strTarUser = oScriptHelper.GetSwitch(arg)

        Case "targetpassword"
            ' Password information
            strTarPassword = oScriptHelper.GetSwitch(arg)
            
        Case "import"
            If (intOperation <> 0) Then
                WScript.Echo L_OnlyOneOper_ErrorMessage
                WScript.Echo L_SeeHelp_Message
                WScript.Quit(ERR_GENERAL_FAILURE)
            End If

            intOperation = OPER_IMPORT

            If oScriptHelper.IsHelpRequested(arg) Then
                DisplayImportHelpMessage
                WScript.Quit(ERR_OK)
            End If
 
        Case "export"
            If (intOperation <> 0) Then
                WScript.Echo L_OnlyOneOper_ErrorMessage
                WScript.Echo L_SeeHelp_Message
                WScript.Quit(ERR_GENERAL_FAILURE)
            End If

            intOperation = OPER_EXPORT

            If oScriptHelper.IsHelpRequested(arg) Then
                DisplayExportHelpMessage
                WScript.Quit(ERR_OK)
            End If

        Case "copy"
            If (intOperation <> 0) Then
                WScript.Echo L_OnlyOneOper_ErrorMessage
                WScript.Echo L_SeeHelp_Message
                WScript.Quit(ERR_GENERAL_FAILURE)
            End If

            intOperation = OPER_COPY

            If oScriptHelper.IsHelpRequested(arg) Then
                DisplayCopyHelpMessage
                WScript.Quit(ERR_OK)
            End If

        Case "save"
            If (intOperation <> 0) Then
                WScript.Echo L_OnlyOneOper_ErrorMessage
                WScript.Echo L_SeeHelp_Message
                WScript.Quit(ERR_GENERAL_FAILURE)
            End If

            intOperation = OPER_SAVE

            If oScriptHelper.IsHelpRequested(arg) Then
                DisplaySaveHelpMessage
                WScript.Quit(ERR_OK)
            End If

        Case "file"
            strFile = oScriptHelper.GetSwitch(arg)

        Case "decpass"
            strDecPass = oScriptHelper.GetSwitch(arg)

        Case "sourcepath"
            strSourcePath = oScriptHelper.GetSwitch(arg)

        Case "destpath"
            strDestPath = oScriptHelper.GetSwitch(arg)

        Case "inherited"
            intFlags = intFlags Or IMPORT_EXPORT_INHERITED
            
        Case "children"
            intFlags = intFlags And Not IMPORT_EXPORT_NODE_ONLY

        Case "merge"
            intFlags = intFlags Or IMPORT_EXPORT_MERGE
            
        Case "enableaudit"
            If (intOperation <> 0) Then
                WScript.Echo L_OnlyOneOper_ErrorMessage
                WScript.Echo L_SeeHelp_Message
                WScript.Quit(ERR_GENERAL_FAILURE)
            End If

            intOperation = OPER_ENABLE_AUDIT

            If oScriptHelper.IsHelpRequested(arg) Then
                DisplayEnableAuditHelpMessage
                WScript.Quit(ERR_OK)
            End If
            
            strMBPath = oScriptHelper.GetSwitch(arg)
        
        Case "disableaudit"
            If (intOperation <> 0) Then
                WScript.Echo L_OnlyOneOper_ErrorMessage
                WScript.Echo L_SeeHelp_Message
                WScript.Quit(ERR_GENERAL_FAILURE)
            End If

            intOperation = OPER_DISABLE_AUDIT

            If oScriptHelper.IsHelpRequested(arg) Then
                DisplayDisableAuditHelpMessage
                WScript.Quit(ERR_OK)
            End If
            
            strMBPath = oScriptHelper.GetSwitch(arg)

        Case "recurse"
            bRecurseAuditOperation = True
            
    End Select
Next

' Check Parameters
If intOperation = 0 Then
    WScript.Echo L_OperationRequired_ErrorMessage
    WScript.Echo L_SeeHelp_Message
    WScript.Quit(ERR_GENERAL_FAILURE)
End If

Select Case intOperation
	Case OPER_COPY
		If strTarServer = "" Or strTarUser = "" Or strTarPassword = "" Then
			WScript.Echo L_RequiredArgsMissing_ErrorMessage
			WScript.Quit(ERR_GENERAL_FAILURE)
		End If
		
	Case OPER_IMPORT
		If strFile = "" Or strSourcePath = "" Or strDestPath = "" Then
			WScript.Echo L_RequiredArgsMissing_ErrorMessage
			WScript.Echo L_SeeImportHelp_Message
			WScript.Quit(ERR_GENERAL_FAILURE)
		End If
	
	Case OPER_EXPORT
		If strFile = "" Or strSourcePath = "" Then
			WScript.Echo L_RequiredArgsMissing_ErrorMessage
			WScript.Echo L_SeeExportHelp_Message
			WScript.Quit(ERR_GENERAL_FAILURE)
		End If
End Select

' Check if /p is specified but /u isn't. In this case, we should bail out with an error
If oScriptHelper.Switches.Exists("password") And Not oScriptHelper.Switches.Exists("user") Then
    WScript.Echo L_PassWithoutUser_ErrorMessage
    WScript.Quit(ERR_GENERAL_FAILURE)
End If

' Check if /u is specified but /p isn't. In this case, we should ask for a password
If oScriptHelper.Switches.Exists("user") And Not oScriptHelper.Switches.Exists("password") Then
    strPassword = oCmdLib.GetPassword
End If

' Initializes authentication with remote machine
intResult = oScriptHelper.InitAuthentication(strServer, strUser, strPassword)
If intResult <> 0 Then
    WScript.Quit(intResult)
End If

' Choose operation
Select Case intOperation
	Case OPER_IMPORT
		intResult = Import(strDecPass, strFile, strSourcePath, strDestPath, intFlags)
		
	Case OPER_EXPORT
		intResult = Export(strDecPass, strFile, strSourcePath, intFlags)

    Case OPER_COPY
        intResult = Repl(strServer, strUser, strPassword, strTarServer, strTarUser, strTarPassword)

    Case OPER_SAVE
        intResult = SaveMD()
        
    Case OPER_ENABLE_AUDIT
		intResult = EnableAudit(strMBPath)
		
	Case OPER_DISABLE_AUDIT
		intResult = DisableAudit(strMBPath) 

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
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General02_Text
    WScript.Echo L_Help_HELP_General03_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General04_Text
    WScript.Echo L_Help_HELP_General05_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General07p2_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Help_HELP_General11_Text
    WScript.Echo L_Help_HELP_General11p1_Text
    WScript.Echo L_Help_HELP_General12_Text
    WScript.Echo L_Help_HELP_General12p1_Text
    WScript.Echo L_Help_HELP_General13_Text
    WScript.Echo L_Help_HELP_General13p1_Text
    WScript.Echo L_Help_HELP_General14_Text
    WScript.Echo L_Help_HELP_General14p1_Text
    WScript.Echo L_Help_HELP_General15_Text
    WScript.Echo L_Help_HELP_General15p1_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General22_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General23_Text
    WScript.Echo L_Help_HELP_General24_Text
    WScript.Echo L_Help_HELP_General25_Text
    WScript.Echo L_Help_HELP_General26_Text
    WScript.Echo L_Help_HELP_General27_Text
    WScript.Echo L_Help_HELP_General28_Text
End Sub

Sub DisplayImportHelpMessage()
    WScript.Echo L_Help_HELP_Import01_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Import02_Text
    WScript.Echo L_Help_HELP_Import02p1_Text
    WScript.Echo L_Help_HELP_Import02p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General07p2_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Help_HELP_Common13_Text
    WScript.Echo L_Help_HELP_Common13p1_Text
    WScript.Echo L_Help_HELP_Common13p2_Text
    WScript.Echo L_Help_HELP_Common14_Text
    WScript.Echo L_Help_HELP_Common15_Text
    WScript.Echo L_Help_HELP_Common15p1_Text
    WScript.Echo L_Help_HELP_Import16_Text
    WScript.Echo L_Help_HELP_Import16p1_Text
    WScript.Echo L_Help_HELP_Import16p2_Text
    WScript.Echo L_Help_HELP_Import16p3_Text
    WScript.Echo L_Help_HELP_Import17_Text
    WScript.Echo L_Help_HELP_Import18_Text
    WScript.Echo L_Help_HELP_Import19_Text
    WScript.Echo L_Help_HELP_Import19p1_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Common21_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Import22_Text
    WScript.Echo L_Help_HELP_Import22p1_Text
End Sub

Sub DisplayCopyHelpMessage()
    WScript.Echo L_Help_HELP_Copy1_Text
    WScript.Echo L_Help_HELP_Copy1p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Copy2_Text
    WScript.Echo L_Help_HELP_Copy2p2_Text
    WScript.Echo L_Help_HELP_Copy2p3_Text
    WScript.Echo L_Help_HELP_Copy3_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Copy4_Text
    WScript.Echo L_Help_HELP_Copy5_Text
    WScript.Echo L_Help_HELP_Copy5p2_Text
    WScript.Echo L_Help_HELP_Copy6_Text
    WScript.Echo L_Help_HELP_Copy7_Text
    WScript.Echo L_Help_HELP_Copy8_Text
    WScript.Echo L_Help_HELP_Copy9_Text
    WScript.Echo L_Help_HELP_Copy10_Text
    WScript.Echo L_Help_HELP_Copy10p2_Text
    WScript.Echo L_Help_HELP_Copy11_Text
    WScript.Echo L_Help_HELP_Copy11p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Copy12_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Copy13_Text
    WScript.Echo L_Help_HELP_Copy14_Text
    WScript.Echo L_Help_HELP_Copy14p2_Text
End Sub

Sub DisplayExportHelpMessage()
    WScript.Echo L_Help_HELP_Export01_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Export02_Text
    WScript.Echo L_Help_HELP_Export02p1_Text
    WScript.Echo L_Help_HELP_Export02p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General07p2_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Help_HELP_Common13_Text
    WScript.Echo L_Help_HELP_Common13p1_Text
    WScript.Echo L_Help_HELP_Common13p2_Text
    WScript.Echo L_Help_HELP_Common14_Text
    WScript.Echo L_Help_HELP_Common15_Text
    WScript.Echo L_Help_HELP_Common15p1_Text
    WScript.Echo L_Help_HELP_Export17_Text
    WScript.Echo L_Help_HELP_Export18_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Common21_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Export22_Text
End Sub

Sub DisplaySaveHelpMessage()
    WScript.Echo L_Help_HELP_Save01_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Save02_Text
    WScript.Echo L_Help_HELP_Save02p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General07p2_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Common21_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Save22_Text
    WScript.Echo L_Help_HELP_Save23_Text
End Sub

Sub DisplayEnableAuditHelpMessage()
    WScript.Echo L_Help_HELP_EnableAudit01_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_EnableAudit02_Text
    WScript.Echo L_Help_HELP_EnableAudit02p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General07p2_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Help_HELP_EnableAudit16_Text
    WScript.Echo L_Help_HELP_EnableAudit16p1_Text
    WScript.Echo L_Help_HELP_EnableAudit17_Text
    WScript.Echo L_Help_HELP_EnableAudit17p1_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Common21_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_EnableAudit22_Text
    WScript.Echo L_Help_HELP_EnableAudit23_Text
    WScript.Echo L_Help_HELP_EnableAudit24_Text
End Sub

Sub DisplayDisableAuditHelpMessage()
    WScript.Echo L_Help_HELP_DisableAudit01_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_DisableAudit02_Text
    WScript.Echo L_Help_HELP_DisableAudit02p2_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_General06_Text
    WScript.Echo L_Help_HELP_General07_Text
    WScript.Echo L_Help_HELP_General07p2_Text
    WScript.Echo L_Help_HELP_General08_Text
    WScript.Echo L_Help_HELP_General09_Text
    WScript.Echo L_Help_HELP_General10_Text
    WScript.Echo L_Help_HELP_DisableAudit16_Text
    WScript.Echo L_Help_HELP_DisableAudit16p1_Text
    WScript.Echo L_Help_HELP_DisableAudit17_Text
    WScript.Echo L_Help_HELP_DisableAudit17p1_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_Common21_Text
    WScript.Echo L_Empty_Text
    WScript.Echo L_Help_HELP_DisableAudit22_Text
    WScript.Echo L_Help_HELP_DisableAudit23_Text
    WScript.Echo L_Help_HELP_DisableAudit24_Text
End Sub

'''''''''''''''''''''''''''
' Copy Function
'''''''''''''''''''''''''''
Function Repl(strSourceServer, strSourceUser, strSourcePwd, strDestServer, strDestUser, strDestPwd)

    If (strSourceServer = ".") Then
        strSourceServer = ""
    End If

    ' Do the first backup

    Dim strBackupCommand
    Dim strSourceDrive, strDrvLetter, strSourcePath
    Dim oShell, oFS, oNetwork
    Dim strDestDrive, strDestPath
    Dim strCopyCommand, strDelCommand, strRestoreCommand

    Set oShell = WScript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        WScript.Echo L_Shell_ErrorMessage
        WScript.Echo L_Shellp2_ErrorMessage
        WScript.Quit(ERR_GENERAL_FAILURE)
    End If

    Set oFS = WScript.CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        WScript.Echo L_FS_ErrorMessage
        WScript.Echo L_FSp2_ErrorMessage
        WScript.Quit(ERR_GENERAL_FAILURE)
    End If

    Set oNetwork = WScript.CreateObject("WScript.Network")
    If Err.Number <> 0 Then
        WScript.Echo L_Network_ErrorMessage
        WScript.Echo L_Networkp2_ErrorMessage
        WScript.Quit(ERR_GENERAL_FAILURE)
    End If

    strBackupCommand =  "cmd /c %SystemRoot%\system32\cscript.exe %SystemRoot%\system32\iisback.vbs /backup"

    If strSourceServer <> "" Then
        strBackupCommand = strBackupCommand & " /s " & strSourceServer 
    Else
        strSourceServer = "127.0.0.1"
    End If

    If strSourceUser <> "" Then
        strBackupCommand = strBackupCommand & " /u " & strSourceUser
    End If

    If strSourcePwd <> "" Then
        strBackupCommand = strBackupCommand & " /p " & strSourcePwd
    End If

    ' need overwrite in case a previous attempt failed
    strBackupCommand = strBackupCommand & " /b iisreplback /overwrite"

    ' backup the source server
    oCmdLib.vbPrintf L_BackingUp_Message, Array(strSourceServer)
    intResult = oShell.Run(strBackupCommand, 1, TRUE)

    WScript.Echo L_BackupComplete_Message

    ' Now map drive to source server
    ' Find a drive letter

    strSourceDrive = "NO DRIVE"
    For strDrvLetter = Asc("C") to Asc("Z")
        If Not oFS.DriveExists(Chr(strDrvLetter)) Then
            strSourceDrive = Chr(strDrvLetter)
            Exit For
        End If
    Next

    If strSourceDrive = "NO DRIVE" Then
        ' No drive letter available
        WScript.Echo L_NoDrive_ErrorMessage
        WScript.Quit(ERR_GENERAL_FAILURE)
    End If

    strSourceDrive = strSourceDrive & ":"
    strSourcePath = "\\" & strSourceServer & "\ADMIN$"

    ' Map the drive
    oCmdLib.vbPrintf L_DriveLetter_Message, Array(strSourceDrive, strSourceServer)

    If strSourceUser <> "" Then
        oNetwork.MapNetworkDrive strSourceDrive, strSourcePath, FALSE, strSourceUser, strSourcePwd
    Else
        oNetwork.MapNetworkDrive strSourceDrive, strSourcePath
    End If

    ' Now map drive to destination server
    ' Find a drive letter

    strDestDrive = "NO DRIVE"
    For strDrvLetter = Asc("C") to Asc("Z")
        If Not oFS.DriveExists(Chr(strDrvLetter)) Then
            strDestDrive = Chr(strDrvLetter)
            Exit For
        End If
    Next

    If strDestDrive = "NO DRIVE" Then
        ' No drive letter available
        WScript.Echo L_NoDrive_ErrorMessage
        WScript.Quit(ERR_GENERAL_FAILURE)
    End If

    strDestDrive = strDestDrive & ":"
    strDestPath = "\\" & strDestServer & "\ADMIN$"

    ' Map the drive
    oCmdLib.vbPrintf L_DriveLetter_Message, Array(strDestDrive, strDestServer)

    If strDestUser <> "" Then
        oNetwork.MapNetworkDrive strDestDrive, strDestPath, FALSE, strDestUser, strDestPwd
    Else
        oNetwork.MapNetworkDrive strDestDrive, strDestPath
    End If

    strCopyCommand = "cmd /c copy /Y " & strSourceDrive & "\system32\inetsrv\metaback\iisreplback.* "
    strCopyCommand = strCopyCommand & strDestDrive & "\system32\inetsrv\metaback"

    ' Copy the files
    WScript.Echo L_Copy_Message
    WScript.Echo strCopyCommand
    intResult = oShell.Run(strCopyCommand, 1, TRUE)
      
    If intResult <> 0 Then
        oCmdLib.vbPrintf L_ReturnVal_ErrorMessage, Array(intResult)
        WScript.Echo L_Copy_ErrorMessage
        WScript.Quit(intResult)
    End If 

    strDelCommand = "cmd /c del /f /q " & strSourceDrive & "\system32\inetsrv\metaback\iisreplback.*"
    intResult = oShell.Run(strDelCommand, 1, TRUE)

    ' Unmap drive to source server
    oCmdLib.vbPrintf L_UnMap_Message, Array(strSourceDrive)
    oNetwork.RemoveNetworkDrive strSourceDrive

    ' Now do the restore on the destination server

    strRestoreCommand = "cmd /c %SystemRoot%\system32\cscript.exe %SystemRoot%\system32\iisback.vbs /restore /s " & strDestServer
    strRestoreCommand = strRestoreCommand & " /u " & strDestUser
    strRestoreCommand = strRestoreCommand & " /p " & strDestPwd
    strRestoreCommand = strRestoreCommand & " /b iisreplback"

    oCmdLib.vbPrintf L_Restoring_Message, Array(strDestServer)
    intResult = oShell.Run(strRestoreCommand, 1, TRUE)

    If intResult <> 0 Then
        oCmdLib.vbPrintf L_ReturnVal_ErrorMessage, Array(intResult)
        WScript.Echo L_Restore_ErrorMessage
        WScript.Quit(intResult)
    End If 

    WScript.Echo L_RestoreComplete_Message

    strDelCommand = "cmd /c del /f /q " & strDestDrive & "\system32\inetsrv\metaback\iisreplback.*"
    intResult = oShell.Run(strDelCommand, 1, TRUE)

    ' Unmap drive to destination server
    oCmdLib.vbPrintf L_UnMap_Message, Array(strDestDrive)
    oNetwork.RemoveNetworkDrive strDestDrive

    WScript.Echo L_CopyComplete_Message
End Function

'''''''''''''''''''''''''''
' Import Function
'''''''''''''''''''''''''''
Function Import(strDecPass, strFile, strSourcePath, strDestPath, intFlags)
    Dim ComputerObj
    Dim strFilePath
    
    On Error Resume Next

    ' Normalize path first
    strFilePath = oScriptHelper.NormalizeFilePath(strFile)
    If Err Then
        Select Case Err.Number
            Case &H80070002
                WScript.Echo L_FileExpected_ErrorMessage
                WScript.Echo L_FileExpectedp2_ErrorMessage

            Case &H80070003
                WScript.Echo L_ParentFolderDoesntExist_ErrorMessage
        End Select
        
        Import = Err.Number
        Exit Function
    End If

    oScriptHelper.WMIConnect
    If Err.Number Then
        WScript.Echo L_WMIConnect_ErrorMessage
        oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        Import = Err.Number
        Exit Function
    End If

    Set ComputerObj = oScriptHelper.ProviderObj.get("IIsComputer='LM'")
    If Err.Number Then
        Select Case Err.Number
            Case &H80070005
                WScript.Echo L_Admin_ErrorMessage
                WScript.Echo L_Adminp2_ErrorMessage

            Case Else
                WScript.Echo L_GetComputerObj_ErrorMessage
                oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        End Select

        Import = Err.Number
        Exit Function
    End If
    
    ' Call Import method from the computer object
    ComputerObj.Import strDecPass, strFilePath, strSourcePath, strDestPath, intFlags
    If Err.Number Then
        Select Case Err.Number
            Case &H80070002
                WScript.Echo L_FileDoesntExist_ErrorMessage

            Case &H8007052B
                WScript.Echo L_IncorrectPassword_ErrorMessage

            Case &H800CC819
                WScript.Echo L_InvalidXML_ErrorMessage

            Case Else
		        WScript.Echo L_Import_ErrorMessage
				WScript.Echo Err.Description
        End Select
        
        Import = Err.Number
        Exit Function
    End If
    
    oCmdLib.vbPrintf L_ConfImported_Text, Array(strSourcePath)
    oCmdLib.vbPrintf L_ConfImportedp2_Text, Array(strFile, strDestPath)
End Function

'''''''''''''''''''''''''''
' Export Function
'''''''''''''''''''''''''''
Function Export(strDecPass, strFile, strSourcePath, intFlags)
    Dim ComputerObj
    Dim strFilePath
    
    On Error Resume Next

    ' Normalize path first
    strFilePath = oScriptHelper.NormalizeFilePath(strFile)
    If Err Then
        Select Case Err.Number
            Case &H80070002
                WScript.Echo L_FileExpected_ErrorMessage
                WScript.Echo L_FileExpectedp2_ErrorMessage

            Case &H80070003
                WScript.Echo L_ParentFolderDoesntExist_ErrorMessage
        End Select
        
        Export = Err.Number
        Exit Function
    End If

    If oScriptHelper.FSObj.FileExists(strFilePath) Then
        WScript.Echo L_FileAlreadyExist_ErrorMessage
        Export = &H80070050
        Exit Function
    End If

    oScriptHelper.WMIConnect
    If Err.Number Then
        WScript.Echo L_WMIConnect_ErrorMessage
        oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        Export = Err.Number
        Exit Function
    End If

    Set ComputerObj = oScriptHelper.ProviderObj.get("IIsComputer='LM'")
    If Err.Number Then
        Select Case Err.Number
            Case &H80070005
                WScript.Echo L_Admin_ErrorMessage
                WScript.Echo L_Adminp2_ErrorMessage

            Case Else
                WScript.Echo L_GetComputerObj_ErrorMessage
                oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        End Select

        Export = Err.Number
        Exit Function
    End If
    
    ' Call Import method from the computer object
    ComputerObj.Export strDecPass, strFilePath, strSourcePath, intFlags
    If Err.Number Then
        WScript.Echo L_Export_ErrorMessage
		WScript.Echo Err.Description
		Export = Err.Number
        Exit Function
    End If

    oCmdLib.vbPrintf L_ConfExported_Text, Array(strSourcePath, strFile)    
End Function

'''''''''''''''''''''''''''
' SaveMD Function
'''''''''''''''''''''''''''
Function SaveMD()
    Dim ComputerObj
    
    On Error Resume Next

    oScriptHelper.WMIConnect
    If Err.Number Then
        WScript.Echo L_WMIConnect_ErrorMessage
        oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        SaveMD = Err.Number
        Exit Function
    End If

    Set ComputerObj = oScriptHelper.ProviderObj.get("IIsComputer='LM'")
    If Err.Number Then
        Select Case Err.Number
            Case &H80070005
                WScript.Echo L_Admin_ErrorMessage
                WScript.Echo L_Adminp2_ErrorMessage

            Case Else
                WScript.Echo L_GetComputerObj_ErrorMessage
                oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
        End Select

        SaveMD = Err.Number
        Exit Function
    End If
    
    ' Call Import method from the computer object
    ComputerObj.SaveData
    If Err.Number Then
        WScript.Echo L_SaveData_ErrorMessage
		WScript.Echo Err.Description
		SaveMD = Err.Number
        Exit Function
    End If

    WScript.Echo L_MDSaved_Text    
End Function

'''''''''''''''''''''''''''
' EnableAudit Function
'''''''''''''''''''''''''''
Function EnableAudit(strMBPath)
    Dim MachineObj
    Dim AclPath, AclPaths
    Dim Result, TempResult
    Dim Index
    
    On Error Resume Next

    Result = 0
    
    If strMBPath = "" Then
        strMBPath = "/"
    End If
    
    ' Add leading slash if necessary
    If Left(strMBPath, 1) <> "/" Then
        strMBPath = "/" & strMBPath
    End If
    
    ' Try to enable audit on the path specified
    Result = EnableAuditInPath(strMBPath)
    If Result <> 0 Then
        DisableAudit = Result
        Exit Function
    End If

    If bRecurseAuditOperation Then
        ' To enable audit recursively, we will need
        ' to find all paths where the AdminACL property is set
        ' under the given path and add a System ACL in each one of them
        Set MachineObj = GetIISObject(strMBPath)
        If Err.Number Then
            oCmdLib.vbPrintf L_GetObject_ErrorMessage, Array("IIS://" & strServer)
            EnableAudit = Err.Number
            Exit Function
        End If
        
        AclPaths = MachineObj.GetDataPaths("AdminACL", IIS_DATA_INHERIT)
        If Err.Number Then
            WScript.Echo L_GetDataPaths_ErrorMessage
            EnableAudit = Err.Number
            Exit Function
        End If

        For Each AclPath In AclPaths
            ' Remove the IIS://<machineName> from the path returned by GetDataPaths
            Index = InStr(7, AclPath, "/")
            If Index > 0 Then
                AclPath = Mid(AclPath, Index)
            Else
                AclPath = "/"
            End If
            
            If NormalizePath(strMBPath) <> NormalizePath(AclPath) Then
                TempResult = EnableAuditInPath(AclPath)
                If TempResult <> 0 Then
                    Result = TempResult
                End If
            End If
        Next
    End If 
    
    EnableAudit = Result
End Function

'''''''''''''''''''''''''''
' EnableAuditInPath Function
'''''''''''''''''''''''''''
Function EnableAuditInPath(strMBPath)
    Dim MBObj
    
    On Error Resume Next

	' Check if auditing is already enabled on that path
	Set MBObj = GetIISObject(strMBPath)
	If Err.Number Then
        oCmdLib.vbPrintf L_GetObject_ErrorMessage, Array(strMBPath)
		EnableAudit = Err.Number
		Exit Function
	End If

	If IsAuditEnabled(MBObj) Then
		oCmdLib.vbPrintf L_AuditAlreadyEnabled_Message, Array(strMBPath)
		EnableAudit = 0
	Else
		AddSystemACE MBObj
		If Err.Number Then
			oCmdLib.vbPrintf L_EnableAuditFail_ErrorMessage, Array(strMBPath)
            oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
            EnableAudit = Err.Number
		Else
			oCmdLib.vbPrintf L_AuditEnabled_Message, Array(strMBPath)
			EnableAudit = 0
		End If
	End If
End Function

'''''''''''''''''''''''''''
' DisableAudit Function
'''''''''''''''''''''''''''
Function DisableAudit(strMBPath)
    Dim MachineObj
    Dim AclPath, AclPaths
    Dim Result, TempResult
    Dim Index
    
    On Error Resume Next

    Result = 0
    
    If strMBPath = "" Then
        strMBPath = "/"
    End If
    
    ' Add leading slash if necessary
    If Left(strMBPath, 1) <> "/" Then
        strMBPath = "/" & strMBPath
    End If

    ' Try to disable audit on the path specified
    Result = DisableAuditInPath(strMBPath)
    If Result <> 0 Then
        DisableAudit = Result
        Exit Function
    End If

    If bRecurseAuditOperation Then
        ' To disable audit recursively, we will need
        ' to find all paths where the AdminACL property is set
        ' under the given path and remove the System ACL in each one of them
        Set MachineObj = GetIISObject(strMBPath)
        If Err.Number Then
            oCmdLib.vbPrintf L_GetObject_ErrorMessage, Array("IIS://" & strServer)
            EnableAudit = Err.Number
            Exit Function
        End If
        
        AclPaths = MachineObj.GetDataPaths("AdminACL", IIS_DATA_INHERIT)
        If Err.Number Then
            WScript.Echo L_GetDataPaths_ErrorMessage
            EnableAudit = Err.Number
            Exit Function
        End If

        For Each AclPath In AclPaths
            ' Remove the IIS://<machineName> from the path returned by GetDataPaths
            Index = InStr(7, AclPath, "/")
            If Index > 0 Then
                AclPath = Mid(AclPath, Index)
            Else
                AclPath = "/"
            End If
            
            If NormalizePath(strMBPath) <> NormalizePath(AclPath) Then
                TempResult = DisableAuditInPath(AclPath)
                If TempResult <> 0 Then
                    Result = TempResult
                End If
            End If
        Next
    End If 
    
    DisableAudit = Result
End Function

'''''''''''''''''''''''''''
' DisableAuditInPath Function
'''''''''''''''''''''''''''
Function DisableAuditInPath(strMBPath)
    Dim MBObj
    
    On Error Resume Next

	' Check if auditing is already enabled on that path
	Set MBObj = GetIISObject(strMBPath)
	If Err.Number Then
        oCmdLib.vbPrintf L_GetObject_ErrorMessage, Array(strMBPath)
		DisableAudit = Err.Number
		Exit Function
	End If

	If Not IsAuditEnabled(MBObj) Then
		oCmdLib.vbPrintf L_AuditAlreadyDisabled_Message, Array(strMBPath)
		DisableAudit = 0
	Else
		' Remove System ACL from security descriptor object
		RemoveSystemACL MBObj
		If Err.Number Then
			oCmdLib.vbPrintf L_DisableAuditFail_ErrorMessage, Array(strMBPath)
            oCmdLib.vbPrintf L_Error_ErrorMessage, Array(Hex(Err.Number), Err.Description)
            DisableAudit = Err.Number
		Else
			oCmdLib.vbPrintf L_AuditDisabled_Message, Array(strMBPath)
			DisableAudit = 0
		End If
	End If
End Function

'''''''''''''''''''''''''''''''
' Add an ACE to the System ACL
'''''''''''''''''''''''''''''''
Sub AddSystemACE(MBObj)
	Dim SecDes, AceObj
	Dim SaclObj, DaclObj
	
	'On Error Resume Next
	
	Set SecDes = MBObj.AdminACL
	Set DaclObj = SecDes.DiscretionaryAcl
	
	Set SaclObj = CreateObject("AccessControlList")
	If Not DaclObj Is Nothing Then
		SaclObj.AclRevision = DaclObj.AclRevision
	Else
		SaclObj.AclRevision = 4
	End If
	
	Set AceObj = CreateObject("AccessControlEntry")
	AceObj.Trustee = ADMINISTRATORS_GROUP_SID
	AceObj.AccessMask = -1 ' Full permissions
	
	SaclObj.AddAce(AceObj)
	
	SecDes.SystemAcl = SaclObj
	MBObj.AdminACL = SecDes
	MBObj.SetInfo
End Sub

''''''''''''''''''''''''''
' Normalizes an IIS Path
''''''''''''''''''''''''''
Function NormalizePath(strPath)
	Dim strIISPath
	Dim strADSIServer
	
	On Error Resume Next
	
	' ADSI uses "localhost" as default server, instead of "." (WMI)
	If strServer = "." Then
		strADSIServer = "localhost"
	Else
		strADSIServer = strServer
	End If
	
	strPath = Replace(strPath, "\", "/")
	
	' Build full IIS path
	strIISPath = "IIS://" & strADSIServer
	If Left(strPath, 1) = "/" Then
		strIISPath = strIISPath & strPath
	Else
		strIISPath = strIISPath & "/" & strPath
	End If

	If Right(strIISPath, 1) = "/" Then
		strIISPath = Left(strIISPath, Len(strIISPath) - 1)
	End If

	NormalizePath = strIISPath	
End Function

''''''''''''''''''''''''''''''''''''
' Retrieves an IIS object via ADSI
'''''''''''''''''''''''''''''''''''
Function GetIISObject(strMBPath)
	Dim ProvObj, IISObj
	Dim strPath

	On Error Resume Next

	strPath = NormalizePath(strMBPath)
	
	If strUser <> "" Then
		Set ProvObj = GetObject("IIS:")
		Set GetIISObject = ProvObj.OpenDsObject( strPath, strUser, strPassword, &H1)
	Else
		Set GetIISObject = GetObject(strPath)
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''
' Removes System ACL from security descriptor
''''''''''''''''''''''''''''''''''''''''''''''
Sub RemoveSystemACL(MBObj)
	Dim SecDes, SaclObj
	
	On Error Resume Next
	
	Set SecDes = MBObj.AdminACL
	SecDes.SystemAcl = Nothing
	MBObj.AdminACL = SecDes
	MBObj.SetInfo
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Verify is audit is already enabled on a certain metabase path
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function IsAuditEnabled(MBObj)
	Dim AdminAcl, SaclObj
	
	On Error Resume Next
	
	Set AdminAcl = MBObj.AdminACL
	Set SaclObj = AdminAcl.SystemAcl

	If SaclObj Is Nothing Then
		IsAuditEnabled = False
	Else
		If Sacl.AceCount > 0 Then
			IsAuditEnabled = True
		Else
			IsAuditEnabled = False
		End If
	End If	
End Function
