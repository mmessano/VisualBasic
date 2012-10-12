' Author: Chrissy LeMaire
' Copyright 2003 NetNerds Consulting Group
' Script is provided AS IS with no warranties or guarantees and assumes no liabilities.
' Website: http://www.netnerds.net

On Error Resume Next 

theDomain = "" ' set this here if you are probing a different domain

Set objNet = CreateObject("WScript.Network")
myFullUsername = trim(objNet.Userdomain & "\" & objNet.UserName)
myComputerName = trim(objNet.ComputerName)

if len(theDomain) = 0 then
theDomain = Trim(objNet.Userdomain)
end if

myUsername = Trim(objNet.UserName)
set objNet = nothing

' in case the groups below do not exist

ADMIN = FALSE
Set objGroup = GetObject("WinNT://" & theDomain & "/Domain Admins")
For Each objUser in objGroup.Members
    If LCase(myusername) = LCase(objUser.Name) Then
     ADMIN = TRUE
    End If
Next

Set objGroup = Nothing

Set objGroup = GetObject("WinNT://" & theDomain & "/Enterprise Admins")
For Each objUser in objGroup.Members
    If LCase(myusername) = LCase(objUser.Name) Then
     ADMIN = TRUE
    End If
Next

Set objGroup = Nothing


IIS = FALSE
err.Clear ' to clear any possible errors from above

  Set objW3SVC = GetObject("IIS://" & myComputerName & "/W3SVC")
  if err.Number = 0 Then
  IIS = TRUE
  end If
  

if ADMIN = FALSE Then
MsgBox "You are not a Domain administrator. This script will not work properly."
Wscript.Quit
End If


if IIS = FALSE Then
MsgBox "IIS is not installed. This script will not work properly."
Wscript.Quit
End If

set objW3SVC = Nothing

on error goto 0

startTime = Timer()


     'Use the WinNT Directory Services
     theDomain = "WinNT://" & theDomain

     'Create the Domain object
     Set objDomain = GetObject(theDomain)

    'Search for Computers in the Domain
     objDomain.Filter = Array("Computer")

On error resume next
serverCount = 0 
IISServerCount = 0 

     For Each Computer In objDomain

 theServer = Computer.Name
 'If lcase(left(theServer,1)) = "s" then 'OPTIONAL; on our network, machines that start with an S are servers
 
 serverCount = ServerCount + 1
 
  
  Set objW3SVC = GetObject("IIS://" & theServer & "/W3SVC")
   
   Select Case Err.Number
   
   ' 70 = permission denied; Strong indication of IIS Server in a DMZ
   ' 462 = The remote server machine does not exist or is unavailable
   ' 429 = ActiveX component can't create object
   ' -2147023169 = CreateObject Failed
   ' -2147012889  = Name could not be resolved
       
  
   case  462, 429
    thehttpVersion = cint(httpVersion(theServer))
    if thehttpVersion > 0 and thehttpVersion < 404 then
    IISServerCount = IISServerCount + 1
    IISTrue = IISTrue & "IIS Server: " & theServer & VbCrLf
    else
    IISFalse = IISFalse & "NOT: " & theServer & VbCrLf
    end if    
   case 70,-2147023169,-2147012889
    IISUnknown = IISUnknown & "Possibly: " & theServer & VbCrLf
   case 0,-2146646000
    IISServerCount = IISServerCount + 1
    IISTrue = IISTrue & "IIS Server: " & theServer & VbCrLf
    set objW3SVC = Nothing
   case else
    'wscript.echo theServer & ": " & errNumber & ", " & err.Description & "<P>" ' for debugging
   end Select
   
  Err.Clear
 'End if

Next

     'Clean up
      Set objDomain = Nothing
  

Function httpVersion(theHost)

On Error Resume Next
    Set objxmlHTTP = createobject("MSXML2.ServerxmlHTTP")

    theURL= "http://" & theHost
    objxmlHTTP.open "GET", theURL, false
    objxmlHTTP.send()
    tempVersion = objxmlHTTP.getResponseHeader("Server")
     If errNumber = -2147012867 Then
      NOSERVER = TRUE
      Else
      NOSERVER = FALSE
     End If
    set objxmlHTTP = nothing
    
    if instr(tempversion,"Microsoft-IIS/") > 0 then
     tempVersion = replace(tempVersion,"Microsoft-IIS/","")
     httpVersion = trim(tempVersion)
    else
     if NOSERVER = TRUE then 
       httpVersion = "404" ' webserver not found ;)
       err.Clear 
     Else 'there was a webserver there, but probably not an IIS Server
      httpVersion = "0"
 End If
    end if
End Function


finishtime = Timer()
totalTime = finishtime - startTime

myStr = myStr &  "Total Time Taken: " & totalTime & " seconds" &  vbCrLf & vbCrLf
myStr = myStr &  "Total Servers Scanned: " & serverCount & vbCrLf
myStr = myStr &  "Total Servers Found: " & IISServerCount &  vbCrLf & vbCrLf
myStr = myStr &  IISTrue & vbCrLf
myStr = myStr &  IISUnknown & vbCrLf
'myStr = myStr &  IISFalse & vbCrLf
myStr = myStr &  "Done" & vbCrLf

wscript.echo myStr