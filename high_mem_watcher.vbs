set events = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecNotificationQuery 
_ 
   ("select * from __instancemodificationevent within 5 where 
targetinstance isa 'Win32_Processor' and targetinstance.LoadPercentage > 50") 
               
if err <> 0 then
   WScript.Echo Err.Description, Err.Number, Err.Source
end if 
' Note this next call will wait indefinitely - a timeout can be specified 
WScript.Echo "Waiting for CPU load events..."
WScript.Echo ""
do 
   set NTEvent = events.nextevent 
   if err <> 0 then
      WScript.Echo Err.Number, Err.Description, Err.Source
      Exit Do
   else      
WScript.Echo NTEvent.TargetInstance.DeviceID WScript.Echo 
NTEvent.TargetInstance.LoadPercentage   
end if
loop

WScript.Echo "finished"
