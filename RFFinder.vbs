Function AddStr(ByRef a, s)
	if IsEmpty(a) Then
		ReDim a(0)
		a(0) = s
	Else
		ReDim Preserve a(UBound(a) + 1)
		a(UBound(a)) = s
	End If
End Function



Function AddStrs(ByRef a1, a2)
	if Not IsEmpty(a2) Then
		Dim	i
		for i = 0 to UBound(a2)
			AddStr a1, a2(i)
		Next
	End If
End Function


Dim	fso
set fso = CreateObject("Scripting.FileSystemObject")



Function GetRFs(dir, mask)
	Dim	result

	' get the directory
	Dim	folder
	On Error Resume Next
	Set folder = fso.GetFolder(dir)
	if Err <> 0 Then
		' no problem if the directory exists... this is even likely
		Exit Function
	End If
	On Error Goto 0

	' get a list of all files in the directory
	Dim	files
	Set files = folder.Files

	' scan for matches
	for each file in files
		if Left(file.Name, Len(mask)) = mask Then
			AddStr result, dir & file.Name
		End If
	Next

	' Done
	GetRFs = result
End Function



Function GetRFCandidates(dirs, client, docID)
	' if we have no directories we of course have no files
	if IsEmpty(dirs) Then
		Exit Function
	End If

	Dim	result

	' if the order ID is 1234, the routing file will start with
	' either "1234." or "1234_"
	mask1 = "" & docID & "."
	mask2 = "" & docID & "_"

	' loop through all the directories
	Dim	i
	for i = 0 to UBound(dirs)
		Dim	dir
		dir = dirs(i)
		if Right(dir, 1) <> "\" Then
			dir = dir & "\"
		End If
		dir = dir & client & "\"

		AddStrs result, GetRFs(dir, mask1)
		AddStrs result, GetRFs(dir, mask2)
	Next

	' done
	GetRFCandidates = result
End Function


' here we create the array of known routing file directories
Dim	RFDirs
AddStr RFDirs, "\\apus\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\ara\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\aries\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\crater\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\europa\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\jupiter\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\leo\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\neptune\Dexma\Docs\RoutingFiles"
AddStr RFDirs, "\\pyxis\Dexma\Docs\RoutingFiles"


' set up the command line arguments
Set objArgs = WScript.Arguments
'WScript.Echo WScript.Arguments.Count
if WScript.Arguments.Count > 2 then Wscript.Echo "Too many arguments."

' get an array of candidates that might be our routing file
Dim	RFCandidates
'				  			 	  		Client Name  Doc ID
RFCandidates = GetRFCandidates(RFDirs, "Wholesale", 15481)

' dump them out
if Not IsEmpty(RFCandidates) Then
	Dim	i
	for i = 0 to UBound(RFCandidates)
		WScript.Echo(RFCandidates(i))
	Next
End If
