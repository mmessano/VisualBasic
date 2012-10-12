' This VBScript code lists the server objects in the site topology.

' ---------------------------------------------------------------
' From the book "Active Directory Cookbook" by Robbie Allen
' Publisher: O'Reilly and Associates
' ISBN: 0-596-00466-4
' Book web site: http://rallenhome.com/books/adcookbook/code.html
' ---------------------------------------------------------------

set objRootDSE = GetObject("LDAP://RootDSE")
strBase    =  "<LDAP://cn=sites," & _
              objRootDSE.Get("ConfigurationNamingContext") & ">;"
strFilter  = "(objectcategory=server);" 
strAttrs   = "distinguishedName;"
strScope   = "subtree"

set objConn = CreateObject("ADODB.Connection")
objConn.Provider = "ADsDSOObject"
objConn.Open "Active Directory Provider"
set objRS = objConn.Execute(strBase & strFilter & strAttrs & strScope)
objRS.MoveFirst
while Not objRS.EOF
    Wscript.Echo objRS.Fields(0).Value
    objRS.MoveNext
wend
