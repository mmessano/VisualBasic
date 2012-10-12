
set dexlog =  CreateObject("Dexma.DexLog")
dexlog.ModuleName = "queue_owner"


' get the computer name
Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputer = WshNetwork.ComputerName


Function gqo(machine, queue)
		set qu = CreateObject("DexConfig.MSMQConfigUtils")
		qu.MachineName = machine
		qu.QueueName = queue
		gqo = qu.GetOwnerAccount
End Function


Function getQueueOwner(machine, queue)
         On Error Resume Next
         owner = gqo(machine, queue)
         if Err = 0 Then
            WScript.Echo("" & machine & "\" & queue & " is owned by " & owner)
            dexlog.Msg("" & machine & "\" & queue & " is owned by " & owner)
         Else
            WScript.Echo("Error querying ownership of " & machine & "\" & queue & ": " & Err.Description)
            dexlog.Msg("Error querying ownership of " & machine & "\" & queue & ": " & Err.Description)
         End If
End Function


Function getOwnerForAllQueuesOnMachine(machine)
	set qs = CreateObject("Dexma.QueueStats")

	set names = qs.GetAllQueueNames(machine)
	for i = 0 to names.Length - 1
	    getQueueOwner machine, names.Item(i)
	next
End Function


' multiple function calls work as well
' just add a call to the function and you're done, such as:
getOwnerForAllQueuesOnMachine strComputer