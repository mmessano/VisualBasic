WScript.Echo "begin test"

serverName = "IAPP520"
fileName   = "\\IAPP520\Dexma\Support\Summit\rso-in-jf.xml"

Set rCdo = CreateObject ("Dexma.RemoteCDO")

If IsNull(rCdo) = vbTrue Then
    WScript.Echo "create 'Dexma.RemoteCDO' failed"
Else
    WScript.Echo "create 'Dexma.RemoteCDO' succeeded"

    rCdo.ServerName = serverName
    rCdo.InputDoc = fileName ' can be file name or XML blob

    rCdo.CreateOrderFromProperties

    WScript.Echo "last error: " & rCdo.LastError
    WScript.Echo "last HRESULT: 0x" & Hex(rCdo.LastHResult)
    WScript.Echo "doc order id: " & rCdo.DocOrderId

    Set rCdo = Nothing
End If

WScript.Echo "test complete"
