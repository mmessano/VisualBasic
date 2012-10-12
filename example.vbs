<package>
   <job id="vbs">
      <script language="VBScript">
         Set WshNetwork = WScript.CreateObject("WScript.Network")
         Set oDrives = WshNetwork.EnumNetworkDrives
         Set oPrinters = WshNetwork.EnumPrinterConnections
         WScript.Echo "Domain = " & WshNetwork.UserDomain
         WScript.Echo "Computer Name = " & WshNetwork.ComputerName
         WScript.Echo "User Name = " & WshNetwork.UserName
         WScript.Echo 
         WScript.Echo "Network drive mappings:"
         For i = 0 to oDrives.Count - 1 Step 2
            WScript.Echo "Drive " & oDrives.Item(i) & " = " & oDrives.Item(i+1)
         Next
         WScript.Echo 
         WScript.Echo "Network printer mappings:"
         For i = 0 to oPrinters.Count - 1 Step 2
            WScript.Echo "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i+1)
         Next
      </script>
   </job>

   <job id="js">
      <script language="JScript">
         var WshNetwork = WScript.CreateObject("WScript.Network");
         var oDrives = WshNetwork.EnumNetworkDrives();
         var oPrinters = WshNetwork.EnumPrinterConnections();
         WScript.Echo("Domain = " + WshNetwork.UserDomain);
         WScript.Echo("Computer Name = " + WshNetwork.ComputerName);
         WScript.Echo("User Name = " + WshNetwork.UserName);
         WScript.Echo();
         WScript.Echo("Network drive mappings:");
         for(i=0; i<oDrives.Count(); i+=2){
            WScript.Echo("Drive " + oDrives.Item(i) + " = " + oDrives.Item(i+1));
         }
         WScript.Echo();
         WScript.Echo("Network printer mappings:");
         for(i=0; i<oPrinters.Count(); i+=2){
            WScript.Echo("Port " + oPrinters.Item(i) + " = " + oPrinters.Item(i+1));
         }
      </script>
   </job>
</package>
