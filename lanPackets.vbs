' Run this script at the command prompt by typing
'     cscript lanPackets.vbs

set oSvc = GetObject("winmgmts:root\cimv2")

wqlQuery = "select BytesReceivedPersec,BytesSentPersec,PacketsReceivedPersec,PacketsSentPersec from Win32_PerfRawData_Tcpip_NetworkInterface where Name = 'Intel[R] PRO_1000 MT Network Connection'"

for each oData in oSvc.ExecQuery(wqlQuery)
	for each oProperty in oData.Properties_
		if oProperty.Name = "PacketsReceivedPersec" then
			packetsReceived = oProperty.Value
		elseif oProperty.Name = "PacketsSentPersec" then
			packetsSent = oProperty.Value
		end if
	next
next

wscript.echo packetsReceived
wscript.echo packetsSent

wscript.echo Date() & " " & Time()

wscript.echo "Packets Received"
wscript.echo "Packets Sent"

