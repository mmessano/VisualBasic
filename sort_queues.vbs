
function MachineName(qPath)
	MachineName = Left(qPath, InStr(qPath, "\") - 1)
end function

function QueueName(qPath)
	QueueName = Mid(qPath, InStr(qPath, "\") + 1)
end function

function GetQueuePaths(opsDSN, client, envt, service)
	' make a temporary filename
	set fileSystem = CreateObject("Scripting.FileSystemObject")
	fileName =  fileSystem.GetSpecialFolder(2).Path & "\" & fileSystem.GetTempName() & ".xml"

	' create the routing file
	set crf = CreateObject("Dexma.CRF")
	crf.CreateRoutingFile opsDSN, client, envt, service, fileName

	' load it as XML
	set rfXML = CreateObject("Microsoft.XMLDOM")
	rfXML.load(fileName)
	fileSystem.DeleteFile(fileName)

	' parse out all the queue names
	set agentElems = rfXML.selectNodes("//agent")
	dim	qPaths()
	redim	qPaths(agentElems.Length)

	agentCount = 0
	for i = 0 to agentElems.Length - 1
		qMachine = agentElems(i).selectSingleNode("target.agent.locale").text
		qName = agentElems(i).selectSingleNode("target.agent.name").text
		if qMachine<>"" and qName<>"" Then
			qPaths(agentCount) = qMachine & "\" & qName
			agentCount = agentCount + 1
		end if
	next
	
	Set oContainer = CreateObject ( "Scripting.Dictionary" )

    j = 0

	for i = 0 to (agentCount - 1)
            If oContainer.Exists(qPaths(i)) = false Then
                oContainer.Add qPaths(i), j
                j = j + 1
            End If
	next

    ' array returned will be unique but not sorted

	GetQueuePaths = oContainer.Keys

End function

WScript.Echo "start"

    aItems = GetQueuePaths ( "opsdb.dexma.com\ops", "EMagic", "PROD", "LP" )

    iStart = LBound(aItems)
    iEnd   = UBound(aItems)
    iCount = iEnd - iStart + 1

    WScript.Echo "aItems has " & iCount & " elements"

    for i = iStart to iEnd
        WScript.Echo "aItems(" & i & ") = " & aItems(i)
    next

WScript.Echo "end"
