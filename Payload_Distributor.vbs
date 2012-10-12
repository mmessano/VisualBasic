Dim SourcePath, SourceFolder
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

SourcePath = InputBox("Enter the UNC path to copy", "Source Path") 
if len(SourcePath) > 0 then

	if right(SourcePath,1) = "\" then
		SourcePath = Left(SourcePath, Len(SourcePath)-1)
	end if
	SourceFolder = Right(SourcePath, Instr(StrReverse(SourcePath), "\")-1)

	Dim Target, MsgTarget
	Target = InputBox("Enter the Environment: " &  chr(13) & "DEVT, QA, IMP, DEMO, PROD", "Destination Machines") 
	if len(Target) > 0 then	
		MsgTarget= MsgBox ("Copy: " &  chr(13) & SourceFolder & chr(13) & "To " & chr(13) & Target & " servers", 1)
		
		if MsgTarget= 1 then
		Dim file
		Dim line
			
		Set file = fso.OpenTextFile( "\\Mensa\Dexma\Data\"& Target & ".txt", 1 )
			
		line = file.ReadLine()

			Do While Not file.AtEndOfStream
				MsgBox "copy to: " & "\\"&line&"\Dexma\Payload\"
				If (fso.FolderExists("\\"&line&"\Dexma\Payload\"& SourceFolder)) Then
					MsgBox " delete " & "\\"&line&"\Dexma\Payload\"& SourceFolder
		     			fso.DeleteFolder "\\"&line&"\Dexma\Payload\"& SourceFolder,1
				End If
  				If (fso.FolderExists("\\"&line&"\Dexma\Payload")) Then		
					fso.CopyFolder SourcePath  , "\\"&line&"\Dexma\Payload\"
				End if
				line = file.ReadLine()
			Loop


		
		MsgBox "Finished!"
	
		end if
	end if
end if	
Set file = Nothing
Set fso = Nothing