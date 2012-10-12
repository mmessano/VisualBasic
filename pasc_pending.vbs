pasc_file = "\\PADoc3\ENT\webSites\Files\PASCFlood\Responses\0711\PASC_24807.xml"

get_description(pasc_file)

Function get_description(arg1)
file = arg1
set oPASCOrderXml = CreateObject("Microsoft.XMLDOM")
oPASCOrderXml.Load(file)
set oNode = oPASCOrderXml.selectSingleNode("//STATUS")
strDescription = oNode.getAttribute("_Description")

set Regex = New RegExp
'Regex.IgnoreCase = True
Regex.Global = True
Regex.Pattern = ".MANUAL PROCESS"

set oMatches = Regex.Execute(strDescription)

WScript.Echo "File: " & pasc_file & vbCRLF
WScript.Echo "Regex: " & Regex.Pattern & vbCRLF
WScript.Echo "Match:" & oMatches(0) & vbCRLF

End Function

'# completed file, no MANUAL status
'&get_description('\\\\PADoc3\\ENT\\webSites\\Files\\PASCFlood\\Responses\\0714\\PASC_24979.xml');
'# outstanding file, contains a MANUAL status
'&get_description('\\\\PADoc3\\MidwestFinancial\\webSites\\Files\\PASCFlood\\Responses\\0703\\PASC_14801.xml');