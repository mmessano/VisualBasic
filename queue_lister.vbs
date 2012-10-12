Function list_all_queues(machine)
		 set qs = CreateObject("Dexma.QueueStats")
		 set names = qs.GetAllQueueNames(machine)
		 for i = 0 to names.Length - 1
		 	 WScript.Echo("'" & machine & "'" & "," & "'" & names.Item(i) & "'")
		 next
End Function

Function list_all_queues2(machine)
		set qs = CreateObject("Dexma.QueueStats")
		set names = qs.GetAllQueueNames(machine)
			Wscript.Echo(vbTab & "<machine name=" & """" & "" & machine & "" & """"& ">")
		for i = 0 to names.Length - 1
			WScript.Echo(vbTab & vbTab & "<Queue name=" & """" & "" & names.Item(i) & "" & """" & "/>")
		next
			WScript.Echo(vbTab & "</machine>")
End Function


Set objArgs = WScript.Arguments
'WScript.Echo WScript.Arguments.Count
WScript.Echo("<?xml version=" & """" & "1.0" & """" & "?>")
WScript.Echo("<servers>")
	For P = 0 to objArgs.Count - 1
'		WScript.Echo objArgs(P)
		list_all_queues2 objArgs(P)
	Next
	WScript.Echo("</servers>")
