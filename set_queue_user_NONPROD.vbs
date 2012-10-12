' This script depends on the MSMQUtilsCOM COM package and the user it is configured to run under
' Typically the COM package runs under dexpromtx for PROD and the corresponding account for other environments
' If that user does not have access to the queue THIS WILL FAIL
' Queues are normally created as part of the DexProcessAdmin service and the queue should be owned by 
'   the user that is designated in the setup dialogue for the listener
' Note that if the designated user is not the same as the user MSMQUtilsCOM runs under THIS WILL FAIL

set dexlog =  CreateObject("Dexma.DexLog")
dexlog.ModuleName = "msmq_users"


' get the computer name
Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputer = WshNetwork.ComputerName

Function squ(machine, queue, user)
		set qu = CreateObject("DexConfig.MSMQConfigUtils")
		qu.MachineName = machine
		qu.QueueName = queue
		qu.GrantAccessToAccount(user)
End Function


Function setQueueUser(machine, queue, user)
         On Error Resume Next
         squ machine, queue, user
         if Err = 0 Then
            WScript.Echo("Gave usership of " & machine & "\" & queue & " to " & user)
            dexlog.Msg("Gave usership of " & machine & "\" & queue & " to " & user)
         Else
            WScript.Echo("Error changing usership of " & machine & "\" & queue & ": " & Err.Description)
            dexlog.Msg("Error changing usership of " & machine & "\" & queue & ": " & Err.Description)
         End If
End Function


Function setUserForAllQueuesOnMachine(machine, user)
	set qs = CreateObject("Dexma.QueueStats")
	set names = qs.GetAllQueueNames(machine)
	for i = 0 to names.Length - 1
	    setQueueUser machine, names.Item(i), user
	next
End Function


' multiple function calls work as well
' just add a call to the function and you're done, such as:
' home_office\Operations - Apps
' home_office\Dex Prod Services  Production group, full permissions
' home_office\Dex Stage Services  Staging group, full permissions
' home_office\Dex Imp Service  Implementation group, full permissions
' home_office\Dex Dev Services  Development group, full permissions
'setUserForAllQueuesOnMachine strComputer, "Everyone"
setUserForAllQueuesOnMachine strComputer , "home_office\New Product Operations"
setUserForAllQueuesOnMachine strComputer , "home_office\Dex Imp Service"

