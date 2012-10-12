Option Explicit

' chris crowe - www.iisfaq.com
' 25 August 2001

' If you are looking for information on IIS then come to www.iisfaq.com
' We have lots of ADSI scripts that may help you in your IIS administration.
'
' We also have 100's of articles, and links to 1000's of articles on IIS
' and related technologies.

dim FSO, IISOBject, IISSite, ServerName, BadPathfound

Function CheckForRootPath(MetabaseObject)
 if (len(MetabaseObject.Path) = 3) then
  WScript.echo "Warning: Metabase Path pointing to root : " & _
   MetabaseObject.AdsPath & " - " & MetabaseObject.Path
 end if
end function

Function ProcessWebSiteVirtualDirectories(MetabasePath)
 Dim VirtualDirectory, WebSiteRoot

' WScript.echo Space(3) & "Scanning " & MetabasePath
 Set WebSiteRoot = getobject(MetabasePath)
 For each VirtualDirectory in WebSiteRoot
  if (VirtualDirectory.Class = "IIsWebVirtualDir") then
   CheckForRootPath(VirtualDirectory)
   if (FSO.FolderExists(VirtualDirectory.Path) = false) then
    wScript.Echo "Invalid VDir Path : " & VirtualDirectory.AdsPath & _
     " - " & VirtualDirectory.Path
    BadPathfound = true
   end if
   ' Scan for child virtual directories
   ProcessWebSiteVirtualDirectories(MetabasePath & "\" & _
    VirtualDirectory.Name)
  end if
 next
 Set WebSiteRoot = Nothing
end function

function SearchWebSite(ServerName, SiteName)
 dim IISWebSite, IISWebSiteRoot, MetabasePath
 
 MetabasePath =  "IIS://" & ServerName & "/w3svc/" & SiteName
 set IISWebSite = GetObject(Metabasepath)
 MetabasePath = MetabasePath & "/Root"
 set IISWebSiteRoot = GetObject(MetabasePath)
 CheckForRootPath(IISWebSiteRoot)
 if (FSO.FolderExists(IISWebSiteRoot.Path) = false) then
  wScript.Echo "Invalid WebSite Path : " & IISWebSiteRoot.AdsPath & _
   " - " & IISWebSiteRoot.Path
  BadPathfound = true
 end if
 Call ProcessWebSiteVirtualDirectories(MetabasePath)
 
 set IISWebSiteRoot = Nothing
 set IISWebSite = nothing
end function

function SearchServer(ServerName)
 set IISOBject = GetObject("IIS://" & ServerName & "/w3svc")
 for each IISSite in IISOBject
  if (IISSite.Class = "IIsWebServer") then
   Call SearchWebSite(ServerName, IISSite.name)
  end if
 next
 set IISOBject = nothing
end function

ServerName = "localhost"

WScript.echo "Searching IIS Metabase for invalid web site & virtual directory paths" & vbcrlf
Set FSO = CreateObject("Scripting.fileSystemObject")
SearchServer(ServerName)
if (BadPathfound = false) then
 WScript.echo "No bad paths were found"
end if 
Set FSO = Nothing