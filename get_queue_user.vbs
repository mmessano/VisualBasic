' This script is NON-FUNCTIONAL at this time.  There is, currently, no method to probe for users of a queue.

Function gqu(machine, queue, user)
		set qu = CreateObject("DexConfig.MSMQConfigUtils")
		qu.MachineName = machine
		qu.QueueName = queue
		gqu = qu.GrantAccessToAccount(user)
End Function


Function setQueueUser(machine, queue, user)
         On Error Resume Next
         squ machine, queue, user
         if Err = 0 Then
            WScript.Echo("Gave usership of " & machine & "\" & queue & " to " & user)
         Else
            WScript.Echo("Error changing usership of " & machine & "\" & queue & ": " & Err.Description)
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

setUserForAllQueuesOnMachine "boise", "Everyone"
'setUserForAllQueuesOnMachine "boise", "home_office\Operations - Apps"
'setUserForAllQueuesOnMachine "boise", "home_office\Dex Stage Services"

