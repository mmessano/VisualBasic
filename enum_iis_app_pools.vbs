Dim locatorObj, ProviderObj, Pools, strQuery

Set locatorObj = CreateObject("WbemScripting.SWbemLocator")
Set ProviderObj = locatorObj.ConnectServer(".", "root/MicrosoftIISv2")

strQuery = "Select * from IIsApplicationPool"
For Each Item In ProviderObj.ExecQuery(strQuery)
    WScript.Echo Replace(Item.Name, "W3SVC/AppPools/", "")
Next
