' This script will count the number of sites you have such as WEB, FTP, SMTP, and NNTP
' 
' For more scripts go to www.iisfaq.com
'
' Chris Crowe - www.iisfaq.com - 23 September 2001

Function CountSites(SiteService, SiteClass, Sitedescription)
Dim IISOBJ, Site, Extra

Sites = 0
on error resume next
Set IISOBJ = GetObject("IIS://localhost/" & SiteService)
for each site in IISOBJ
	if (Site.Class = SiteClass) then
		Sites = sites + 1
	wscript.echo Site.servercomment
	end if
next
if (sites = 0) then
  Sites = "no"
  Extra = "s"
elseif (Sites = 1) then
  Extra = ""
else  
  Extra = "s"
end if
WScript.echo "You have " & Sites & " " & SiteDescription & " site" & Extra & "."
Set IISOBJ = nothing
CountSites = Sites
end Function

call CountSites("W3SVC",    "IIsWebServer", "WEB")
call CountSites("MSFTPSVC", "IIsFtpServer", "FTP")
call CountSites("SMTPSVC",  "IIsSmtpServer", "SMTP")
call CountSites("NNTPSVC",  "IIsNntpServer", "NNTP")
