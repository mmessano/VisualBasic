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
		 for i = 0 to names.Length - 1
		 	 WScript.Echo("" & names.Item(i) & "")
		 next
End Function

Function list_all_queues3(machine)
		 set qs = CreateObject("Dexma.QueueStats")
		 set names = qs.GetAllQueueNames(machine)
		 for i = 0 to names.Length - 1
		 	 WScript.Echo("" & machine & "" & "/" & "" & names.Item(i) & "")
		 next
End Function


Set objArgs = WScript.Arguments
'WScript.Echo WScript.Arguments.Count
For P = 0 to objArgs.Count - 1
'  WScript.Echo objArgs(P)
list_all_queues objArgs(P)
Next
