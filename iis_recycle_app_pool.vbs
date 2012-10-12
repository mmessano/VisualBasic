strServer = "MyServer" //Replace this with the target server name

strUser = "Administrator" //Provide Administrator privilege credentials

strPassword = "password" //Input Administrator privileged account password


Set LocatorObj = CreateObject("WBemScripting.SWbemLocator")

LocatorObj.Security_.ImpersonationLevel = 3        Impersonate

LocatorObj.Security_.AuthenticationLevel = 6        Pkt Privacy 
(required for remote administration over WMI as of Win2k3 SP1)

Set ProviderObj = LocatorObj.ConnectServer(strServer, "root/MicrosoftIISv2", strUser, strPassword)

Set MyAppPool = ProviderObj.Get( "IIsApplicationPool=w3svc/apppools/DefaultAppPool" 
MyAppPool.Recycle