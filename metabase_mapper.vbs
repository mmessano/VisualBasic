' metabase_mapper.vbs
' queries the local IIS metabase and returns all objects installed
Class IISMetaBase
	Private fso, f, sFilePath

	Public Sub MapMetaBase(byVal filePath)
		sFilePath = filePath
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(filePath, 8, True)

		GetThisIISPath "IIS://LocalHost", 0

		f.close
		Set f = Nothing
		Set fso = Nothing
	End Sub

	Private Function WriteTab(ByVal tabSize)
		Dim s

		s = ""
		if tabSize > 0 then
			For i = 1 to tabSize
				s = s & vbTab
			Next
		end if

		WriteTab = s
	End Function

	Private Sub GetThisIISPath(byVal ADsPath, byval tabSize)
		Dim oIIS, sItem, i

		Set oIIS = GetObject(ADsPath)

		' Provide a list of all objects installed on this machine
		For Each sItem in oIIS
			With sItem
				f.Write WriteTab(tabSize) & "Name:" & _
					vbtab & vbtab & .Name & vbcrlf
				f.Write WriteTab(tabSize) & "Class:" & _
					vbtab & vbtab & .Class & vbcrlf
				f.Write WriteTab(tabSize) & "ADsPath:" & _
					vbtab & .ADsPath & vbcrlf
				f.Write WriteTab(tabSize) & "GUID:" & _
					vbtab & vbtab & .GUID & vbcrlf
				f.Write WriteTab(tabSize) & "Parent:" & _
					vbtab & vbtab & .Parent & vbcrlf
				f.Write WriteTab(tabSize) & "Schema:" & _
					vbtab & vbtab & .Schema & vbcrlf & vbcrlf

				'recurse any sub nodes...
				GetThisIISPath .ADsPath, tabSize + 1
			End With
		Next

		Set oIIS = Nothing
	End Sub

	Private Sub Class_Terminate()
		MsgBox "IIS Metabase Mapped to " & sFilePath
	End Sub
End Class



Dim oMap

Set oMap = New IISMetaBase
oMap.MapMetaBase "c:\Dexma\out\IIS_Setup.txt"
Set oMap = Nothing