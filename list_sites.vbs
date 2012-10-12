' This script will list the number of sites you have such as WEB, FTP, SMTP, and NNTP
' 
' For more scripts go to www.iisfaq.com
'
' Chris Crowe - www.iisfaq.com - 23 September 2001

Function ListSites(SiteService, SiteClass, Sitedescription)
Dim IISOBJ, Site

Sites = 0
on error resume next
Set IISOBJ = GetObject("IIS://localhost/" & SiteService)
for each site in IISOBJ
	if (Site.Class = SiteClass) then
		wscript.echo Site.servercomment
	end if
next

'WScript.echo "You have " & Sites & " " & SiteDescription & " site" & Extra & "."
Set IISOBJ = nothing
end Function



call ListSites("W3SVC",    "IISWEBsERVER", "WEB")
call ListSites("MSFTPSVC", "IISFTPSERVER", "FTP")
call ListSites("SMTPSVC",  "IISSMTPSERVER", "SMTP")
call ListSites("NNTPSVC",  "IISNNTPSERVER", "NNTP")

Function ListSites2(Server)
Dim IISOBJ

end Function
