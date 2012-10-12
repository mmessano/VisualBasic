Option Explicit

' global vars
Dim m_oLogger, m_sBinDir 
Set m_oLogger = createobject("Dexma.DexLog")
m_oLogger.ModuleName = "PayloadCleanUp_Log"

m_sBinDir = ""	' used to prepend to component file names to create the full path.


wscript.echo("Beginning payload clean up...")
m_oLogger.Msg("******************************************************************")
m_oLogger.Msg("Beginning payload clean up...")

Call Main("")	' pass the path to look for the config files, eventually a CMD line param

wscript.echo(" ")
m_oLogger.Msg("Payload clean up complete")
m_oLogger.Msg("******************************************************************")
wscript.echo("Payload clean up complete, check log file for details : PayloadCleanUp_Log")


Sub Main(sConfigPath)

	' shutdown DexProcessController
	Dim oWshShell, sResult
	Set oWshShell = CreateObject("WScript.Shell")
	
	wscript.echo("MAIN: Stopping DexProcessController...")
	m_oLogger.Msg("MAIN: Stopping DexProcessController...")
	sResult = oWshShell.Run ("net stop DexProcessController", 1, true)
	if ( sResult <> 0 ) then	' error
		Dim sContinue
		m_oLogger.Msg("MAIN: DexProcessController was NOT stopped, waiting for user input.")
		WScript.StdOut.Write("DexProcessController was NOT stopped, would you like to continue?  Type 'Yes' or 'No' then press the ENTER key: ")
		WScript.StdIn.Read(0)
		sContinue = WScript.StdIn.ReadLine()
		if(LCase(sContinue) <> "yes") then
			wscript.echo("MAIN: Cleanup process stopped by user")
			m_oLogger.Msg("MAIN: Cleanup process stopped by user")
			wscript.quit
		end if
		m_oLogger.Msg("MAIN: User chose to continue cleanup process")
	else
		m_oLogger.Msg("MAIN: Stopped DexProcessController...")
		wscript.echo("MAIN: Stopped DexProcessController...")
	end if	' if ( sResult <> 0 ) then

	
	' parse through the server list and decide what server to "clean"
	Dim sServerName
	Dim oComponentXml, oMTXComponents, oNTComponents, oMTXComponent, oNTComponent
	Dim sComponentName, sComponentType, sComponentsFilePath, sComponentFileName
	Dim sRegKeyName, sRegKeyPath, sInput
	Dim sProcessCtrlCfgPath

	' retrieve the server name
	Dim oWshNetwork
	Set oWshNetwork = CreateObject("WScript.Network")
	sServerName = oWshNetwork.ComputerName 
	
	' load the components xml and determine what components to clean
	m_oLogger.Msg("MAIN: Loading component config file: " & sConfigPath & "PayloadCleanUp_ComponentList.xml")
	Set oComponentXml = createobject("Microsoft.XMLDOM")
	oComponentXml.Load(sConfigPath & "PayloadCleanUp_ComponentList.xml")
	
	' get the root dir where components are installed (should be e:\dexma\bin)
	sComponentsFilePath = oComponentXml.SelectSingleNode("//binDir").text
	m_sBinDir = sComponentsFilePath
	
	' get the path to the process controller config file
	sProcessCtrlCfgPath = oComponentXml.SelectSingleNode("//processControllerConfig").text
	
	' get a list of all the mts components
	Set oMTXComponents = oComponentXml.SelectNodes("//components/component[@installType='mts']")
	
	' get a list of all the NT service components
	Set oNTComponents = oComponentXml.SelectNodes("//components/component[@installType='ntservice']")

		m_oLogger.Msg("******** server start ***************")
		m_oLogger.Msg("MAIN: Attempting to clean server: " & sServerName)
		m_oLogger.Msg("MAIN: Bin directory: " & sComponentsFilePath)
		m_oLogger.Msg("MAIN: Process Controller Config Path directory: " & sProcessCtrlCfgPath)
		
		
		' loop through the mts components
		wscript.echo("MAIN: Removing MTX(COM+) packages...")
		m_oLogger.Msg("MAIN: Removing MTX(COM+) packages...")
		for each oMTXComponent in oMTXComponents
			' get all the info from the node
			sRegKeyName = oMTXComponent.GetAttribute("regKey")
			sRegKeyPath = oMTXComponent.GetAttribute("regKeyPath")
			if ( sRegKeyPath = "" ) then	' default to "Software\Dexma\"
				sRegKeyPath = "Software\Dexma\"
			end if	'if ( sRegKeyPath = "" ) then
			sComponentType = oMTXComponent.GetAttribute("installType")
			sComponentFileName = oMTXComponent.GetAttribute("fileName")
			sComponentName  = getComponentName(oMTXComponent)
			
			' log and clean
			wscript.echo("MAIN: Attempting to remove component: " & sComponentName & " : " & sComponentType & " from " & sServerName)
			m_oLogger.Msg("MAIN: Attempting to remove component: " & sComponentName & " : " & sComponentType & " from " & sServerName)

			Call DeleteMTXPackage(sComponentName)

			' delete the registry key
			if ( sRegKeyName <> "") then	' attempt to delete the reg key
				Call DeleteRegKey(sServerName, sRegKeyName, sRegKeyPath)
			else
				m_oLogger.Msg("MAIN: Registry key attribute is empty, no attempt to delete keys for component:" & sComponentName)
				wscript.echo("MAIN: Registry key attribute is empty for component:" & sComponentName)
			end if	'if ( sRegKeyName <> "") then
		next	' for each oMTXComponent in oMTXComponents
		wscript.echo("MAIN: Done removing MTX(COM+) packages...")
		m_oLogger.Msg("MAIN: Done removing MTX(COM+) packages...")


		' pause between "milestones" wait for the user to press "enter"
		wscript.echo(" ")
		wscript.echo("Next step: remove NT services")
		Call pauseForUserInput()
		
		' loop throuh the NT service components
		wscript.echo("Removing NT Services...")
		m_oLogger.Msg("MAIN: Removing NT Services packages...")
		
		for each oNTComponent in oNTComponents
			' get all the info from the node
			sRegKeyName = oNTComponent.GetAttribute("regKey")
			sRegKeyPath = oNTComponent.GetAttribute("regKeyPath")
			if ( sRegKeyPath = "" ) then	' default to "Software\Dexma\"
				sRegKeyPath = "Software\Dexma\"
			end if	'if ( sRegKeyPath = "" ) then
			sComponentType = oNTComponent.GetAttribute("installType")
			sComponentFileName = oNTComponent.GetAttribute("fileName")
			sComponentName  = getComponentName(oNTComponent)
			
			' log and clean
			wscript.echo("MAIN: Attempting to remove component: " & sComponentName & " : " & sComponentType & " from " & sServerName)
			m_oLogger.Msg("MAIN: Attempting to remove component: " & sComponentName & " : " & sComponentType & " from " & sServerName)

			' check to see if this service is a LOCAL service or DCOM'd to another server.
			' if it is DCOM'd, re-register it BEFORE trying to unregister it so it is completely removed from the registry and DCOM.
			Dim oWMIService, oColItems, oItem
			
			Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")	' the local computer
			Set oColItems = oWMIService.ExecQuery("Select * from Win32_DCOMApplicationSetting")
			For Each oItem in oColItems

				if( lcase(oItem.Caption) = lcase(sComponentName) ) then	
					if ( trim(oItem.LocalService) = "" or IsNull(oItem.LocalService) ) then	' the service is NOT local, assume DCOM'd to other server and re-register
						Call RegisterNTComponent(sComponentName, sComponentsFilePath)
					end if	' if ( trim(oItem.LocalService) = "" or IsNull(oItem.LocalService) ) then
				end if	' if( oItem.Caption = sComponentName ) then
			Next	' end For Each oItem in oColItems
			
			Call DeleteNTService(sComponentName, sComponentsFilePath)

			' delete the registry key
			if ( sRegKeyName <> "") then	' attempt to delete the reg key
				Call DeleteRegKey(sServerName, sRegKeyName, sRegKeyPath)
			else
				m_oLogger.Msg("MAIN: Registry key attribute is empty, no attempt to delete keys for component:" & sComponentName)
				wscript.echo("MAIN: Registry key attribute is empty for component:" & sComponentName)
			end if	'if ( sRegKeyName <> "") then
		next	' for each oNTComponent in oNTComponents
		wscript.echo("MAIN: Done removing NT Services...")
		m_oLogger.Msg("MAIN: Done removing NT Services packages...")


		' pause between "milestones" wait for the user to press "enter"
		wscript.echo(" ")
		wscript.echo("Next step:  Unregister files using 'regsvr32 -u -s' (.dll, .wsc, etc)")
		Call pauseForUserInput()

		' get the list of files to delete/unregister
		'Dim oFiles
		'Set oFiles = oComponentXml.SelectNodes("//files/file")
		
		' unregister DLL and WSC files
		wscript.echo("MAIN: Unregistering files...")
		m_oLogger.Msg("MAIN: Unregistering files...")
		'Call UnregFiles(oFiles)
		Call UnregFiles(oComponentXml)
		m_oLogger.Msg("MAIN: Done Unregistering files...")
		wscript.echo("MAIN: Done Unregistering files...")
		
		' pause between "milestones" wait for the user to press "enter"
		wscript.echo(" ")
		wscript.echo("Next step: Delete component files (.dll, .exe, etc)")
		Call pauseForUserInput()

		' delete the files off the server
		wscript.echo("MAIN: Deleting files...")
		m_oLogger.Msg("MAIN: Deleting files...")
		'Call DeleteFiles(oFiles)
		Call DeleteFiles(oComponentXml)
		m_oLogger.Msg("MAIN: Done deleting files...")
		wscript.echo("MAIN: Done deleting files...")
		
		' pause between "milestones" wait for the user to press "enter"
		wscript.echo(" ")
		wscript.echo("Next step:  Delete folders specified in the configuration xml")
		Call pauseForUserInput()
		
		' delete the directories off the server
		wscript.echo("MAIN: Deleting folders...")
		m_oLogger.Msg("MAIN: Deleting folders...")
		Dim oFolders
		Set oFolders= oComponentXml.SelectNodes("//folders/folder")
		Call DeleteFolders(oFolders)
		m_oLogger.Msg("MAIN: Done deleting folders...")
		wscript.echo("MAIN: Done deleting folders...")
		
		' pause between "milestones" wait for the user to press "enter"
		wscript.echo(" ")
		wscript.echo("Next step:  Delete elements from the process controller (DexProcessController) configuration xml")
		Call pauseForUserInput()	
	
		' clean up the process controller config file
		wscript.echo("MAIN: Deleting process controller elements...")
		m_oLogger.Msg("MAIN: Deleting process controller elements...")
		Call CleanProcessController( sProcessCtrlCfgPath, oComponentXml)
		m_oLogger.Msg("MAIN: Done deleting process controller elements...")
		wscript.echo("MAIN: Done deleting process controller elements...")

		' log the end of this server and pause
		m_oLogger.Msg("MAIN: Server cleaned: " & sServerName )
		m_oLogger.Msg("******** server end ***************")
	
		' restart DexProcessAdmin
			sResult = ""
			wscript.echo(" ")
			wscript.echo("MAIN: Restarting DexProcessController...")
			m_oLogger.Msg("MAIN: Restarting DexProcessController...")
			sResult = oWshShell.Run ("net start DexProcessController", 1, true)
			if ( sResult <> 0 ) then	' error
				m_oLogger.Msg("MAIN: ERROR: DexProcessController was NOT restarted.")
				WScript.echo("MAIN:  ERROR: DexProcessController was NOT restarted.  Please check the service manually")
			else
				m_oLogger.Msg("MAIN: Restarted DexProcessController...")
				wscript.echo("MAIN: Restarted DexProcessController...")
			end if	' if ( sResult <> 0 ) then

		' tell the user we are done
		wscript.echo(" ")
		WScript.Echo "MAIN: Done cleaning server: " & sServerName


	' clean up
	Set oMTXComponents	= Nothing
	Set oNTComponents	= Nothing
	Set oMTXComponent	= Nothing
	Set oNTComponent	= Nothing
	Set oComponentXml	= Nothing
	Set oWshNetwork		= Nothing
	Set oWshShell		= Nothing

End Sub	' Main(sConfigPath)

'************** functions/subs **************

Sub DeleteRegKey(sComputer, sKeyName, sKeyPath)
	
	Dim oRegistry, sKeyToDelete, Return
	sKeyToDelete = sKeyPath & sKeyName	' sKeyPath will be an input to the function
	
	wscript.echo("DeleteRegKey: Attempting to delete : " & sComputer & "\HKEY_LOCAL_MACHINE\" & sKeyToDelete)
	m_oLogger.Msg("DeleteRegKey: Attempting to delete : " & sComputer & "\HKEY_LOCAL_MACHINE\" & sKeyToDelete)
	
	const HKEY_LOCAL_MACHINE = &H80000002
	Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\default:StdRegProv")

	' check for children keys and remove them
	Dim aSubKeys, sSubKey
	On Error Resume Next
	oRegistry.EnumKey HKEY_LOCAL_MACHINE, sKeyToDelete, aSubKeys	' enumerate subkeys

	if ( IsArray(aSubKeys) ) Then
		For Each sSubKey In aSubKeys	' check the children for subkeys
			wscript.echo("DeleteRegKey: Found child key : " & sSubKey)
			m_oLogger.Msg("DeleteRegKey: Found child key : " & sSubKey)
			Call DeleteRegKey(sComputer, "\" & sSubKey , sKeyToDelete)
			wscript.echo("DeleteRegKey: Deleted child key : " & sSubKey)
			m_oLogger.Msg("DeleteRegKey: Deleted child key : " & sSubKey)
		Next	' end For Each sSubKey In aSubKeys
	end if
	
	'Delete key
		Return = oRegistry.DeleteKey(HKEY_LOCAL_MACHINE, sKeyToDelete)
		If (Return = 0) And (Err.Number = 0) Then    
			wscript.echo("DeleteRegKey: Key deleted : " & sComputer & "\HKEY_LOCAL_MACHINE\" & sKeyToDelete)
			m_oLogger.Msg("DeleteRegKey: Key deleted : " & sComputer & "\HKEY_LOCAL_MACHINE\" & sKeyToDelete)
		Else
			m_oLogger.Msg("DeleteRegKey: WARN: Delete Key failed - key may not exist. Error = " & Err.Number & "  Key: " & sComputer & "\HKEY_LOCAL_MACHINE\" & sKeyToDelete)
			wscript.echo("DeleteRegKey: WARN: Delete Key failed - key may not exist. Error = " & Err.Number & "  Key: " & sComputer & "\HKEY_LOCAL_MACHINE\" & sKeyToDelete)
		End If

	Set oRegistry = Nothing

End Sub	' DeleteRegKey(sCOmputer, sKeyName, sKeyPath)

Sub DeleteMTXPackage(sPackageToDelete)

	Dim oCatalog, oPackages, iPackageCount, iCounter, bPkgExists
	bPkgExists = False
	Set oCatalog = CreateObject("MTSAdmin.Catalog.1")

	Set oPackages = oCatalog.GetCollection("Packages")
	oPackages.Populate

	iPackageCount = oPackages.Count

	wscript.echo("DeleteMTXPackage: Attempting to delete package : " & sPackageToDelete)
	m_oLogger.Msg("DeleteMTXPackage: Attempting to delete package : " & sPackageToDelete)

	For iCounter = iPackageCount - 1 to 0 step -1
		if lCase(oPackages.Item(iCounter).Name) = LCase(sPackageToDelete) Then
			bPkgExists = true
			oPackages.Remove(iCounter)
				if err <> 0 Then
					wscript.echo("DeleteMTXPackage: ERROR: Error occurred deleting packaging: " & err.Description )
					m_oLogger.Msg("DeleteMTXPackage: ERROR: Error occurred deleting packaging: " & err.Description )
				else
					oPackages.SaveChanges
					if err <> 0 then
						wscript.echo("DeleteMTXPackage: ERROR: Error occurred saving changes: " & err.Description)
						m_oLogger.Msg("DeleteMTXPackage: ERROR: Error occurred saving changes: " & err.Description)
					else
						wscript.echo( "DeleteMTXPackage: " & sPackageToDelete & " package deleted successfully")
						m_oLogger.Msg( "DeleteMTXPackage: " & sPackageToDelete & " package deleted successfully")
					end if	' if err <> 0 then
					
				End If	' if err <> 0 Then
				
		end if	' if oPackages.Item(i).Name = pkgArray(x) Then
	Next	' For i = iPackageCount - 1 to 0 step -1
	
	if ( not bPkgExists) then	' package not found on the server, log it.
		wscript.echo( "DeleteMTXPackage: WARN: " & sPackageToDelete & " not found on server")
		m_oLogger.Msg( "DeleteMTXPackage: WARN: " & sPackageToDelete & " not found on server")
	end if
	
	' clean up
	set oPackages = nothing
	set oCatalog = nothing

End Sub	' DeleteMTXPackage(sPackageToDelete)

Sub DeleteNTService(sServiceToDelete, sServiceFilePath)
	
	Dim oWshShell, sResult
	Set oWshShell = CreateObject("WScript.Shell")
	
	wscript.echo("DeleteNTService: Stopping service: " & sServiceToDelete)
	m_oLogger.Msg("DeleteNTService: Stopping service: " & sServiceToDelete)
	sResult = oWshShell.Run ("net stop " & sServiceToDelete, 1, true)
	m_oLogger.Msg(sResult)

	wscript.echo("DeleteNTService: Unregistering service: " & sServiceFilePath & sServiceToDelete)
	m_oLogger.Msg("DeleteNTService: Unregistering service: " & sServiceFilePath & sServiceToDelete)
	m_oLogger.Msg("DeleteNTService: Command code to be run: " & sServiceFilePath & sServiceToDelete & " -unregserver")
	On Error Resume Next
	sResult = oWshShell.Run (sServiceFilePath & sServiceToDelete & " -unregserver", 1, true)
	if Err.number <> 0 then	' error
		wscript.echo("DeleteNTService: ERROR: Error unregistering service: " & sServiceToDelete & " : " & sResult)
		m_oLogger.Msg("DeleteNTService: ERROR: Error unregistering service: " & sServiceToDelete & " : " & sResult)
	else
		wscript.echo("DeleteNTService: " & sServiceToDelete & " service unregistered")
		m_oLogger.Msg("DeleteNTService: " & sServiceToDelete & " service unregistered")
	end if
	On Error GoTo 0
		
	Set oWshShell = Nothing
End Sub	' DeleteNTService(sServiceToDelete, sServiceFilePath)

Sub RegisterNTComponent(sComponentName, sComponentsFilePath)
	Dim oWshShell, sResult
	Set oWshShell = CreateObject("WScript.Shell")
	
	wscript.echo("RegisterNTComponent: Stopping service: " & sComponentName)
	m_oLogger.Msg("RegisterNTComponent: Stopping service: " & sComponentName)
	sResult = oWshShell.Run ("net stop " & sComponentName, 1, true)
	m_oLogger.Msg(sResult)

	wscript.echo("RegisterNTComponent: Re-registering service: " & sComponentsFilePath & sComponentName)
	m_oLogger.Msg("RegisterNTComponent: Re-registering service: " & sComponentsFilePath & sComponentName)
	On Error Resume Next
	sResult = oWshShell.Run (sComponentsFilePath & sComponentName & " -regserver", 1, true)
	if Err.number <> 0 then	' error
		wscript.echo("RegisterNTComponent: WARNING: Failed re-registering component: " & sComponentName & " : " & sResult)
		m_oLogger.Msg("RegisterNTComponent: WARNING: Failed re-registering component: " & sComponentName & " : " & sResult)
	else
		wscript.echo("RegisterNTComponent: " & sComponentName & " component re-registered")
		m_oLogger.Msg("RegisterNTComponent: " & sComponentName & " component re-registered")
	end if
	On Error GoTo 0
		
	Set oWshShell = Nothing
End Sub	' RegisterNTComponent(sComponentName, sComponentsFilePath)

Sub DeleteFiles(ByVal oComponentXML)
	' for each oFile in oFiles delete the file (path attribute) from the server
	Dim oFile, oFs, oFiles, sFullFilePath
	Dim oComponents, oComponent, sFullCompFilePath
	
	Set oFs = CreateObject("Scripting.FileSystemObject")
	
	' loop through the component files listed
	Set oComponents = oComponentXml.SelectNodes("//components/component")
	
	For Each oComponent in oComponents
		sFullCompFilePath = m_sBinDir & oComponent.GetAttribute("fileName")
		
		wscript.echo("DeleteFiles: Attempting to delete file: " & sFullCompFilePath)
		m_oLogger.Msg("DeleteFiles: Attempting to delete file: " & sFullCompFilePath)
		if ( oFs.FileExists(sFullCompFilePath) ) then	' delete it
			On Error Resume Next
			Call oFs.DeleteFile(sFullCompFilePath, true)	' delete, forcing read-only deletes, too
			if Err.number <> 0 then	' error
				wscript.echo("DeleteFiles: ERROR: Error deleting file: " & sFullCompFilePath & " : " & Err.Description)
				m_oLogger.Msg("DeleteFiles: ERROR: Error deleting file: " & sFullCompFilePath & " : " & Err.Description)
			else
				wscript.echo("DeleteFiles: File deleted: " & sFullCompFilePath)
				m_oLogger.Msg("DeleteFiles: File deleted: " & sFullCompFilePath)
			end if
			On Error Goto 0
			
		else	' log that it did not exist
			wscript.echo("DeleteFiles: WARN: File does not exist: " & sFullCompFilePath)
			m_oLogger.Msg("DeleteFiles: WARN: File does not exist: " & sFullCompFilePath)
		end if	' if ( oFs.FileExists(sFileUNC) ) then
	Next	' For Each oComponent in oComponents
		
	
	' loop through the additional files listed
	Set oFiles = oComponentXml.SelectNodes("//files/file")
	
	For Each oFile in oFiles
		sFullFilePath = oFile.GetAttribute("path")
		
		wscript.echo("DeleteFiles: Attempting to delete file: " & sFullFilePath)
		m_oLogger.Msg("DeleteFiles: Attempting to delete file: " & sFullFilePath)
		if ( oFs.FileExists(sFullFilePath) ) then	' delete it
			On Error Resume Next
			Call oFs.DeleteFile(sFullFilePath, true)	' delete, forcing read-only deletes, too
			if Err.number <> 0 then	' error
				wscript.echo("DeleteFiles: ERROR: Error deleting file: " & sFullFilePath & " : " & Err.Description)
				m_oLogger.Msg("DeleteFiles: ERROR: Error deleting file: " & sFullFilePath & " : " & Err.Description)
			else
				wscript.echo("DeleteFiles: File deleted: " & sFullFilePath)
				m_oLogger.Msg("DeleteFiles: File deleted: " & sFullFilePath)
			end if
			On Error Goto 0
			
		else	' log that it did not exist
			wscript.echo("DeleteFiles: WARN: File does not exist: " & sFullFilePath)
			m_oLogger.Msg("DeleteFiles: WARN: File does not exist: " & sFullFilePath)
		end if	' if ( oFs.FileExists(sFileUNC) ) then
	Next	' For Each oFile in oFiles

	Set oFs		= Nothing
	Set oComponent	=	Nothing
	Set oComponents =	Nothing
	Set oFile		=	Nothing	
	Set oFiles		=	Nothing

End Sub	' Sub DeleteFiles(oFiles)

Sub DeleteFolders(oFolders)
	' for each oFolder in oFolders delete the folder (path attribute) from the server
	
	Dim oFolder, oFs
	Dim sFullFolderPath
	Set oFs = CreateObject("Scripting.FileSystemObject")
	
	For Each oFolder in oFolders
		sFullFolderPath = oFolder.GetAttribute("path")
		
		wscript.echo("DeleteFolders: Attempting to delete folder: " & sFullFolderPath)
		m_oLogger.Msg("DeleteFolders: Attempting to delete folder: " & sFullFolderPath)
		if ( oFs.FolderExists(sFullFolderPath) ) then	' delete it
			On Error Resume Next
			Call oFs.DeleteFolder(sFullFolderPath, true)	' delete, forcing read-only deletes, too
			if Err.number <> 0 then	' error
				wscript.echo("DeleteFolders: ERROR: Error deleting folder: " & sFullFolderPath & " : " & Err.Description)
				m_oLogger.Msg("DeleteFolders: ERROR: Error deleting folder: " & sFullFolderPath & " : " & Err.Description)
			else
				wscript.echo("DeleteFolders: folder deleted: " & sFullFolderPath)
				m_oLogger.Msg("DeleteFolders: folder deleted: " & sFullFolderPath)
			end if
			On Error Goto 0
			
		else	' log that it did not exist
			wscript.echo("DeleteFolders: WARN: folder does not exist: " & sFullFolderPath)
			m_oLogger.Msg("DeleteFolders: WARN: folder does not exist: " & sFullFolderPath)
		end if	' if ( oFs.FolderExists(sFullFolderPath) ) then
	Next	' For Each oFolder in oFolders

	Set oFs		= Nothing
	Set oFolder	= Nothing
	
End Sub	' Sub DeleteFolders(oFolders)


Sub unregSvr32(sFullDllPath)


End Sub	' unregSvr32(sFullDllPath)

Sub UnregFiles(ByVal oComponentXml)
	' this routine will loop through all the components and files in the XML and 
	' perform a regsvr32 -u on the sFilePath file if the file extension is .dll or.wsc

	Dim oWshShell, sResult
	Dim oFiles, oFile, sFilePath, sFileExt
	Dim oComponents, oComponent, sCompFilePath

	Set oWshShell = CreateObject("WScript.Shell")

	Set oComponents = oComponentXml.SelectNodes("//components/component")	
	' loop through the components	
	For Each oComponent in oComponents
		' get the file path and check the extension
		sCompFilePath = m_sBinDir & oComponent.GetAttribute("fileName")
		sFileExt = LCASE(right(sCompFilePath, 3))
		
		if ( sFileExt = "wsc" or sFileExt = "dll" ) then	' run the regsvr32 -u command
			wscript.echo("UnregFile: Unregistering file: " & sCompFilePath)
			m_oLogger.Msg("UnregFile: Unregistering file: " & sCompFilePath)
			On Error Resume Next
			sResult = oWshShell.Run ("regsvr32 -u -s " & sCompFilePath , 1, true)
			if Err.number <> 0 then	' error
				wscript.echo("UnregFile: ERROR: Error unregistering file: " & sCompFilePath & " : " & sResult)
				m_oLogger.Msg("UnregFile: ERROR: Error unregistering file: " & sCompFilePath & " : " & sResult)
			else
				wscript.echo("UnregFile: " & sCompFilePath & " file unregistered")
				m_oLogger.Msg("UnregFile: " & sCompFilePath & " file unregistered")
			end if
			On Error GoTo 0
		end if	' if ( sFileExt = "wsc" or sFileExt = "dll" ) then
	
	Next	' For Each oFile in oFiles
	
	
	' reset the var
	sFileExt = ""
	
	Set oFiles = oComponentXml.SelectNodes("//files/file")
	
	' loop through the files
	For Each oFile in oFiles
		' get the file path and check the extension
		sFilePath = oFile.GetAttribute("path")
		sFileExt = LCASE(right(sFilePath, 3))
		
		if ( sFileExt = "wsc" or sFileExt = "dll" or sFileExt = "exe" ) then	' run the regsvr32 -u command
			wscript.echo("UnregFile: Unregistering file: " & sFilePath)
			m_oLogger.Msg("UnregFile: Unregistering file: " & sFilePath)
			On Error Resume Next
			sResult = oWshShell.Run ("regsvr32 -u -s " & sFilePath , 1, true)
			if Err.number <> 0 then	' error
				wscript.echo("UnregFile: ERROR: Error unregistering file: " & sFilePath & " : " & sResult)
				m_oLogger.Msg("UnregFile: ERROR: Error unregistering file: " & sFilePath & " : " & sResult)
			else
				wscript.echo("UnregFile: " & sFilePath & " file unregistered")
				m_oLogger.Msg("UnregFile: " & sFilePath & " file unregistered")
			end if
			On Error GoTo 0
		end if	' if ( sFileExt = "wsc" or sFileExt = "dll" ) then
	
	Next	' For Each oFile in oFiles
	
	Set oComponent	=	Nothing
	Set oComponents =	Nothing
	Set oFile		=	Nothing	
	Set oFiles		=	Nothing
	Set oWshShell	=	Nothing
		
End Sub	' Sub UnregFiles(ByVal oFiles)


Sub CleanProcessController( sProcessCtrlCfgPath, ByVal oComponentXml )
	' this routine will remove elements from the process controller config file.
	Dim oProcessXML, oProcesses, oProcess, oComponents, oComponent, sComponentName, oParent
	Set oProcessXML = createobject("Microsoft.XMLDOM")
	oProcessXML.Load(sProcessCtrlCfgPath)

	wscript.echo("CleanProcessController: Cleaning configuration file " & sProcessCtrlCfgPath)
	m_oLogger.Msg("CleanProcessController: Cleaning configuration file " & sProcessCtrlCfgPath)

	Set oComponents = oComponentXml.SelectNodes("//components/component")
	
	' loop through the components, if there is a matching process in the process controller config, delete it.
	For Each oComponent in oComponents
		'sComponentName = oComponent.GetAttribute("name")
		sComponentName = getComponentName(oComponent)
		wscript.echo("CleanProcessController: Searching for: " & sComponentName)
		m_oLogger.Msg("CleanProcessController: Searching for: " & sComponentName)

		Set oProcesses = oProcessXML.SelectNodes("//process[@name='" & sComponentName & "']")
		
		if ( oProcesses.Length = 0 ) then
			wscript.echo("CleanProcessController: WARN: Error removing configuration for: " & sComponentName & " : entry not found in configuration file")
			m_oLogger.Msg("CleanProcessController: WARN: Error removing configuration for: " & sComponentName & " : entry not found in configuration file")
		else	' remove the node from the XML
			wscript.echo("CleanProcessController: Found processes: " & sComponentName)
			m_oLogger.Msg("CleanProcessController: Found processes: " & sComponentName)
			For Each oProcess in oProcesses
				' remove each of the nodes
				Set oParent = oProcess.parentNode
				oParent.removeChild(oProcess)
				wscript.echo("CleanProcessController: Removed node(s) from configuration: " & sComponentName )
				m_oLogger.Msg("CleanProcessController: Removed node(s) from configuration: " & sComponentName )
			Next	'For Each oProcess in oProcesses
			Set oProcess = Nothing		
		end if	'if ( IsNull(oProcess) ) then
		
	Next	' For Each oComponent in oComponents
	wscript.echo("CleanProcessController: Done removing nodes from configuration file: " & sProcessCtrlCfgPath )
	m_oLogger.Msg("CleanProcessController: Done removing nodes from configuration file: " & sProcessCtrlCfgPath )

	oProcessXML.save(sProcessCtrlCfgPath)

	Set oProcess	=	Nothing
	Set oProcesses	=	Nothing
	Set oComponent	=	Nothing
	Set oComponents	=	Nothing
	Set oProcessXML	=	Nothing
	Set oParent		=	Nothing
End Sub	' Sub CleanProcessController( sProcessCtrlCfgPath, oComponentXml )

Function getComponentName(oComponent)
	' takes a component node and returns the value of the installName attribute if it exists or returns
	' the name to the left of the "." of the fileName attribute.
	Dim sComponentName, sComponentFileName, iDotIndex
	sComponentFileName = oComponent.GetAttribute("fileName")
	sComponentName = oComponent.GetAttribute("installName")
		if ( sComponentName  = "" ) then	' there was no install name
			iDotIndex = InStr( sComponentFileName, "." )
			sComponentName = Left( sComponentFileName, iDotIndex - 1 )
		end if
	
	getComponentName = sComponentName
	
End Function	' getComponentName(oComponent)

Sub pauseForUserInput()
	Dim sInput, stdIn, stdOut
	sInput = ""
	Set stdOut = Wscript.StdOut
	Set stdIn = WScript.StdIn

	
	StdOut.Write "Press the ENTER key to continue.  Type 'quit' to terminate program "
	StdIn.Read(0)
	sInput = WScript.StdIn.ReadLine()
	if ( lcase(sInput) = "quit" ) then
		wscript.echo("Cleanup process stopped by user")
		m_oLogger.Msg("Cleanup process stopped by user")
		wscript.quit
	end if
	
	wscript.echo(" ")
	Set stdIn	= Nothing
	Set stdOut	= Nothing
	
End Sub	' Sub pauseForUserInput()