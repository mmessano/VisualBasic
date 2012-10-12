strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_PerfRawData_PerfOS_Processor Where Name = '0'")
For Each objItem in colItems
    CounterValue1 = objItem.PercentUserTime
    TimeValue1 = objItem.TimeStamp_Sys100NS
Next
For i = 1 to 5
    Wscript.Sleep(1000)
    Set colItems = objWMIService.ExecQuery _
       ("Select * From Win32_PerfRawData_PerfOS_Processor Where Name = '0'")
    For Each objItem in colItems
        CounterValue2 = objItem.PercentUserTime
        TimeValue2 = objItem.TimeStamp_Sys100NS
        If TimeValue2 - TimeValue1 = 0 Then
            Wscript.Echo "Percent User Time = 0%"
        Else
            PercentProcessorTime = 100 * (CounterValue2 - CounterValue1) / _
                (TimeValue2 - TimeValue1)
            Wscript.Echo "Percent User Time = " & _
                PercentProcessorTime & "%"
        End if
        CounterValue1 = CounterValue2
        TimeValue1 = TimeValue2
    Next
Next