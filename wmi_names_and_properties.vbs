strComputer = "."
strNamespace = "root\cimv2"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
For Each objclass2 in objWMIService.SubclassesOf()
    If Left(objClass2.Path_.Class,13) = "Win32_PerfRaw" Then
        strClass = objClass2.Path_.Class     
        Set objClass = GetObject("winmgmts:\\" & strComputer & _
            "\" & strNameSpace & ":" & strClass)
        For Each objClassProperty In objClass.Properties_
            strType = ""
            strFormula = ""
            For Each objClassQualifier In objClassProperty.Qualifiers_
                If objClassQualifier.Name = "countertype" Then
                Select Case objClassQualifier.Value
                    Case 0 
                        strType = "PERF_COUNTER_RAWCOUNT_HEX"
                    Case 1073874176
                        strType = "PERF_AVERAGE_BULK"
                    Case 1073939457
                        strType = "PERF_SAMPLE_BASE"
                    Case 1073939458
                        strType = "PERF_AVERAGE_BASE"
                    Case 1073939459
                        strType = "PERF_RAW_BASE"
                    Case 1073939712
                        strType = "PERF_LARGE_RAW_BASE"
                    Case 1107494144
                        strType = "PER_COUNTER_MULTI_BASE"
                    Case 256
                        strType = "PERF_COUNTER_LARGE_RAWCOUNT_HEX"
                    Case 272696320
                        strType = "PERF_COUNTER_COUNTER"
                    Case 272696576
                        strType = "PERF_COUNTER_BULK_COUNT"
                    Case 2816
                        strType = "PERF_COUNTER_TEXT"
                    Case 591463680
                        strType = "PERF_COUNTER_MULTI_TIMER_INV"
                    Case 4195238
                        strType = "PERF_COUNTER_DELTA"
                    Case 4195584
                        strType = "PERF_COUNTER_LARGE_DELTA"
                    Case 4260864
                        strType = "PERF_SAMPLE_COUNTER"
                    Case 4523008
                        strType = "PERF_COUNTER_QUELEN_TYPE"
                    Case 537003008
                        strType = "PERF_RAW_FRACTION"
                    Case 541525248
                        strType = "PERF_PRECISION_SYSTEM_TIMER"
                    Case 558957824
                        strType = "PERF_100NSEC_TIMER_INV"
                    Case 542180608
                        strType = "PERF_100NSEC_TIMER"
                    Case 542573824
                        strType = "PERF_PRECISION_100NS_TIMER"
                    Case 543229184
                        strType = "PERF_OBJ_TIME_TIMER"
                    Case 549585920
                        strType = "PERF_SAMPLE_FRACTION"
                    Case 4523264
                        strType = "PERF_COUNTER_LARGE_QUELEN_TYPE"
                    Case 5571840
                        strType = "PERF_COUNTER_100NS_QUELEN_TYPE"
                    Case 541132032
                        strType = "PERF_COUNTER_TIMER"
                    Case 574686464
                        strType = "PERF_COUNTER_MULTI_TIMER"
                    Case 575735040
                        strType = "PERF_100NSEC_MULTI_TIMER"
                    Case 592512256
                        strType = "PERF_100NSEC_MULTI_TIMER_INV"
                    Case 65536
                        strType = "PERF_COUNTER_RAWCOUNT"
                    Case 65792
                        strType = "PERF_COUNTER_LARGE_RAWCOUNT"
                    Case 6620416
                        strType = "PERF_COUNTER_OBJ_TIME_QUELEN_TYPE"
                    Case 805438464
                        strType = "PERF_AVERAGE_TIMER"
                    Case 807666944
                        strType = "PERF_ELAPSED_TIME"
                    Case 557909248
                        strType = "PERF_COUNTER_TIMER_INV"
                    Case Else
                        strType = "Counter type could not be determined."
                    End Select
                End If
            Next
            WScript.Echo strClass & "." & objClassProperty.Name & "," & strType
        Next
    End If
Next