' Run this script at the command prompt by typing
'     cscript lanTraffic.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select BytesReceivedPersec,BytesSentPersec,PacketsReceivedPersec,PacketsSentPersec from Win32_PerfRawData_Tcpip_NetworkInterface where Name = 'Intel[R] PRO_1000 MT Network Connection'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "BytesReceivedPersec" then
			bytesReceived = oProperty.Value
		elseif oProperty.Name = "BytesSentPersec" then
			bytesSent = oProperty.Value
		end if
	next
next

wscript.echo bytesReceived
wscript.echo bytesSent

wscript.echo Date() & " " & Time()

wscript.echo "Bytes Received"
wscript.echo "Bytes Sent"

