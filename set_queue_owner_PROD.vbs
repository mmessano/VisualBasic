' This script depends on the MSMQUtilsCOM COM package and the user it is configured to run under
' Typically the COM package runs under dexpromtx for PROD and the corresponding account for other environments
' If that user does not have access to the queue THIS WILL FAIL
' Queues are normally created as part of the DexProcessAdmin service and the queue should be owned by 
'   the user that is designated in the setup dialogue for the listener
' Note that if the designated user is not the same as the user MSMQUtilsCOM runs under THIS WILL FAIL

set dexlog =  CreateObject("Dexma.DexLog")
dexlog.ModuleName = "msmq_owner"


Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputer = WshNetwork.ComputerName

Function sqo(machine, queue, owner)
		set qu = CreateObject("DexConfig.MSMQConfigUtils")
		qu.MachineName = machine
		qu.QueueName = queue
		qu.SetOwnerAccount(owner)
End Function


Function setQueueOwner(machine, queue, owner)
         On Error Resume Next
         sqo machine, queue, owner
         if Err = 0 Then
            WScript.Echo("Gave ownership of " & machine & "\" & queue & " to " & owner)
            dexlog.Msg("Gave ownership of " & machine & "\" & queue & " to " & owner)
         Else
            WScript.Echo("Error changing ownership of " & machine & "\" & queue & ": " & Err.Description)
            dexlog.Msg("Error changing ownership of " & machine & "\" & queue & ": " & Err.Description)
         End If
End Function


Function setOwnerForAllQueuesOnMachine(machine, owner)
	set qs = CreateObject("Dexma.QueueStats")
	set names = qs.GetAllQueueNames(machine)
	for i = 0 to names.Length - 1
	    setQueueOwner machine, names.Item(i), owner
	next
End Function


' multiple function calls work as well
' just add a call to the function and you're done, such as:
'  Production
'setOwnerForAllQueuesOnMachine "apollo", "home_office\dexpromtx"
'  Staging
'setOwnerForAllQueuesOnMachine "boise", "home_office\dexstagemtx"
setOwnerForAllQueuesOnMachine strComputer, "home_office\dexpromtx"
