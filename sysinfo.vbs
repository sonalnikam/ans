set oShell= CreateObject("WScript.Shell")
strComputer = oShell.ExpandEnvironmentStrings("%ComputerName%")

Const ForAppending = 8
Const ForWriting = 2


Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("c:\" & strComputer & ".html", ForWriting, True)

objLogFile.write "<html>"


objLogFile.write "<head>"
objLogFile.write "<meta http-equiv=Content-Type content=""text/html; charset=windows-1252"">"
objLogFile.write "<meta name=Generator content=""Microsoft Word 11 (filtered)"">"
objLogFile.write "<title>" & strComputer & "</title>"
objLogFile.write "<style>"
objLogFile.write "<!--"
objLogFile.write " /* Style Definitions */"
objLogFile.write " p.MsoNormal, li.MsoNormal, div.MsoNormal"
objLogFile.write "	{margin:0in;"
objLogFile.write "	margin-bottom:.0001pt;"
objLogFile.write "	font-size:12.0pt;"
objLogFile.write "	font-family:""Times New Roman"";}"
objLogFile.write "@page Section1"
objLogFile.write "	{size:8.5in 11.0in;"
objLogFile.write "	margin:1.0in 1.25in 1.0in 1.25in;}"
objLogFile.write "div.Section1"
objLogFile.write "	{page:Section1;}"
objLogFile.write "-->"
objLogFile.write "</style>"

objLogFile.write "</head>"

objLogFile.write "<body lang=EN-US>"

objLogFile.write "<div class=Section1>"

    objLogFile.write "<body>"

    objLogFile.write("<p class=MsoNormal><b>" & "Installation Report for " & StrComputer & " at " & Date() & " " & Time() & "</b></p>")

    objLogFile.write("<p class=MsoNormal><b>" & "Created with Sysinfo.vbs 2.1" & "</b></p>")
    objLogFile.write("<p class=MsoNormal>" & "___________________________" & "</p>")
    objLogFile.write("<p class=MsoNormal><b>" & "Hardware Information" & "</b></p>")

Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
    objLogFile.write("<p class=MsoNormal>" & "System Name: " & objComputer.Name & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "System Manufacturer: " & objComputer.Manufacturer & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "System Model: " & objComputer.Model & "</p>")
Next

Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
For each objBIOS in colBIOS
    objLogFile.write("<p class=MsoNormal>" & "Serial Number: " & objBIOS.SerialNumber & "</p>")
Next

Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
    objLogFile.write("<p class=MsoNormal>" & "Total Physical Memory: " & round(objComputer.TotalPhysicalMemory/1073741824) & " GB</p>")
Next


Set colSettings = objWMIService.ExecQuery("Select * from Win32_Processor")
x = 0
For Each objProcessor in colSettings 
x=x+1
ProcDesc=objProcessor.Description
Next
    objLogFile.write("<p class=MsoNormal>" & "Processor: " & x & " X [" & ProcDesc & "]</p>")

'--------------------------------------------------------------------------------------
    objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b >" & "OS Information" & "</b></p>")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colSettings 
    OS = split(objOperatingSystem.Name,"|")
    objLogFile.write("<p class=MsoNormal>" & "OS Name: " & "<font color=""#A43900""><b>" & OS(0) & "</b></font color></p>") 
    objLogFile.write("<p class=MsoNormal>" & "SystemRoot: " & "<font color=""#A43900""><b>"  & OS(1) & "</b></font color></p>") 
    objLogFile.write("<p class=MsoNormal>" & "System Volume: " & "<font color=""#A43900""><b>"  & OS(2) & "</b></font color></p>") 
    objLogFile.write("<p class=MsoNormal>" & "Version: " & objOperatingSystem.Version & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "Service Pack: " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "Windows Directory: " & objOperatingSystem.WindowsDirectory & "</p>")
Next

Set colItems = objWMIService.ExecQuery("Select * from Win32_TimeZone")
For Each objItem in colItems
    objLogFile.write("<p class=MsoNormal>" & "TimeZone Bias: " & objItem.Bias & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "TimeZone: " & objItem.Caption & "</p>")
Next


'--------------------------------------------------------------------------------------
    objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b>" & "Hotfix Information" & "</b></p>")
Set colQuickFixes = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering")
objLogFile.write("<table border=""1"">")
For Each objQuickFix in colQuickFixes
if objQuickFix.Description <> "" then
objLogFile.write("<tr><td>")
    objLogFile.write("<p class=MsoNormal>" & "Description: " & objQuickFix.Description & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "Hot Fix ID: " & objQuickFix.HotFixID & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "Installation Date: " & objQuickFix.InstallDate & "</p>")
    objLogFile.write("<p class=MsoNormal>" & "Installed By: " & objQuickFix.InstalledBy & "</p>")
objLogFile.write("</td></tr>")
end if
Next
objLogFile.write("</table>")
 

'--------------------------------------------------------------------------------------
    objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b>" & "Network Information" & "</b></p>")
Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True") 
 
n = 1

objLogFile.write("<table border=""1"">")
 For Each objAdapter in colAdapters
  objLogFile.write("<tr><td>")
     objLogFile.Write("<p class=MsoNormal><b>" &  "Network Adapter " & n & "</b></p>")
     objLogFile.Write("<p class=MsoNormal>" &  "  Description: " & objAdapter.Description & "</p>")
     objLogFile.Write("<p class=MsoNormal>" &  "  Physical (MAC) address: " & objAdapter.MACAddress & "</p>")
 
   If Not IsNull(objAdapter.IPAddress) Then
      For i = 0 To UBound(objAdapter.IPAddress)
             objLogFile.Write("<p class=MsoNormal>" &  "IP address: " & objAdapter.IPAddress(i) & "</p>")
      Next
   End If
 
   If Not IsNull(objAdapter.IPSubnet) Then
      For i = 0 To UBound(objAdapter.IPSubnet)
             objLogFile.Write("<p class=MsoNormal>" &  "Subnet Mask: " & objAdapter.IPSubnet(i) & "</p>")
     Next
   End If
 
   If Not IsNull(objAdapter.DefaultIPGateway) Then
      For i = 0 To UBound(objAdapter.DefaultIPGateway)
             objLogFile.Write("<p class=MsoNormal>" &  "Default gateway: " & objAdapter.DefaultIPGateway(i) & "</p>")
      Next
   End If
 
       objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
       objLogFile.write("<p class=MsoNormal><b>" & "Name Servers" & "</b></p>")
       objLogFile.write("<p class=MsoNormal>" & "DNS servers in search order:" & "</p>")
 
   If Not IsNull(objAdapter.DNSServerSearchOrder) Then
      For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
             objLogFile.Write("<p class=MsoNormal>" & objAdapter.DNSServerSearchOrder(i) & "</p>")
      Next
   End If
 
       objLogFile.Write("<p class=MsoNormal>" & "DNS domain: " & objAdapter.DNSDomain & "</p>")

   If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
      For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
             objLogFile.Write("<p class=MsoNormal>" & "    DNS suffix search list: " & objAdapter.DNSDomainSuffixSearchOrder(i) & "</p>")
      Next
   End If
 
       objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
       objLogFile.Write("<p class=MsoNormal><b>" & "DHCP" & "</b></p>")
       objLogFile.Write("<p class=MsoNormal>" & "DHCP enabled: " & objAdapter.DHCPEnabled & "</p>")

       objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
  if not isnull(objAdapter.WINSPrimaryServer) then
       objLogFile.write("<p class=MsoNormal><b>" & "WINS" & "</b></p>")
       objLogFile.write("<p class=MsoNormal>" & "Primary WINS server:   " & objAdapter.WINSPrimaryServer & "</p>")
       objLogFile.write("<p class=MsoNormal>" & "Secondary WINS server: " & objAdapter.WINSSecondaryServer & "</p>") 
  end if
  n = n + 1
objLogFile.write("</td></tr>")
 Next
objLogFile.write("</table>")
 
'--------------------------------------------------------------------------------------
Function WMIDateStringToDate(utcDate)
   WMIDateStringToDate = CDate(Mid(utcDate, 5, 2)  & "/" & _
                               Mid(utcDate, 7, 2)  & "/" & _
                               Left(utcDate, 4)    & " " & _
                               Mid (utcDate, 9, 2) & ":" & _
                               Mid(utcDate, 11, 2) & ":" & _
                               Mid(utcDate, 13, 2))
End Function
'--------------------------------------------------------------------------------------
objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
objLogFile.write("<p class=MsoNormal><b>" & "Service Information" & "</b></p>")
objLogFile.write("<table border=""1"">")
objLogFile.write("<tr>")
objLogFile.Write("<td>Display Name</td>"  _
& "<td>Service Name</td>" _
& "<td>Service Type</td>"  _
& "<td>Service State</td>"  _ 
& "<td>Service Started</td>"  _
& "<td>Start Mode</td>"  _
& "<td>Account Name</td>") 
objLogFile.Write("</tr>")

Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService in colListOfServices
objLogFile.Write("<tr>")
    objLogFile.Write("<td>" & objService.DisplayName) & "</td>" 
    objLogFile.Write("<td>" & objService.Name) & "</td>" 
    objLogFile.Write("<td>" & objService.ServiceType) & "</td>" 
    objLogFile.Write("<td>" & objService.State) & "</td>" 
    objLogFile.Write("<td>" & objService.Started) & "</td>" 
    objLogFile.Write("<td>" & objService.StartMode) & "</td>" 
    objLogFile.Write("<td>" & objService.StartName) & "</td>" 
objLogFile.Write("</tr>")
Next
    objLogFile.Write("</table>")


'--------------------------------------------------------------------------------------
    objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b>" & "Disk Information" & "</b></p>")
objLogFile.write("<table border=""1"">")
objLogFile.write("<tr>")
objLogFile.Write _
    ("<td>Description" & "</td>" _  
& "<td>DeviceID" & "</td>" _  
& "<td>DriveType" & "</td>" _ 
& "<td>FileSystem" & "</td>" _  
& "<td>Size" & "</td>" _ 
& "<td>FreeSpace" & "</td>" _  
& "<td>Space Used" & "</td>" _ 
& "<td>MediaType" & "</td>" _ 
& "<td>Name" & "</td>" _ 
& "<td>VolumeName" & "</td>" _  
& "<td>VolumeSerialNumber" & "</td>") 
objLogFile.write("</tr>")

Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
For each objDisk in colDisks
objLogFile.write("<tr>")
    objLogFile.Write "<td>" & objDisk.Description & "</td>"   
    objLogFile.Write "<td>" & objDisk.DeviceID & "</td>" 
    objLogFile.Write "<td>" & objDisk.DriveType & "</td>" 
    objLogFile.Write "<td>" & objDisk.FileSystem & "</td>" 
  If objDisk.Description = "Local Fixed Disk" then
    objLogFile.Write "<td>" & round(objDisk.Size/1073741824) & "</td>" 
    objLogFile.Write "<td>" & round(objDisk.FreeSpace/1073741824) & "</td>" 
    objLogFile.Write "<td>" & round((objDisk.Size-objDisk.FreeSpace)/1073741824) & "</td>" 
  End If
    objLogFile.Write "<td>" & objDisk.MediaType & "</td>" 
    objLogFile.Write "<td>" & objDisk.Name & "</td>" 
    objLogFile.Write "<td>" & objDisk.VolumeName & "</td>" 
    objLogFile.Write "<td>" & objDisk.VolumeSerialNumber & "</td>" 
objLogFile.write("</tr>")
Next
objLogFile.write("</table>")


'--------------------------------------------------------------------------------------
    objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b>" & "User/Group Information" & "</b></p>")
Set colGroups = GetObject("WinNT://" & strComputer & "")
colGroups.Filter = Array("group")
objLogFile.write("<table border=""1"">")

For Each objGroup In colGroups
	objLogFile.write("<tr><td>")
        objLogFile.write("<p class=MsoNormal><b>Group Name: " & objGroup.Name & "</b></p>")
    For Each objUser in objGroup.Members
        objLogFile.Write("<p class=MsoNormal><i>" & objUser.Name & "</i></p>")
    Next
	objLogFile.write("</td></tr>")
Next
objLogFile.write("</table>")


'--------------------------------------------------------------------------------------
    objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b>" & "Shares" & "</b></p>")
Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
objLogFile.write("<table border=""1"">")
objLogFile.write("<tr>")
objLogFile.write("<td>" & "Caption</td>")
objLogFile.write("<td>" & "Name</td>")
objLogFile.write("<td>" & "Path</td>")
objLogFile.write("</tr>")

For each objShare in colShares
objLogFile.write("<tr>")
    objLogFile.Write("<td>" & objShare.Caption & "</td>") 
    objLogFile.Write("<td>" & objShare.Name & "</td>")
    objLogFile.Write("<td>" & objShare.Path & "</td>")   
objLogFile.write("</tr>")
Next
objLogFile.write("</table>")	
'----------------------------------------------------------------------------
	objLogFile.write "<p class=MsoNormal>&nbsp;</p>"
    objLogFile.write("<p class=MsoNormal><b>" & "Network Adapters/Speed" & "</b></p>")
 Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM MSNdis_LinkSpeed",,48)
objLogFile.write("<table border=""1"">")
objLogFile.write("<tr>")
objLogFile.write("<td>" & "Network Adapter</td>")
objLogFile.write("<td>" & "Link Speed</td>")
objLogFile.write("</tr>")

For Each objItem in colItems
if not mid(objItem.InstanceName,1,8)="WAN Mini" then
objLogFile.write("<tr>")
    objLogFile.Write("<td>" & objItem.InstanceName & "</td>") 
    objLogFile.Write("<td>" & objItem.NdisLinkSpeed/10000000 & " Gbps" & "</td>")
objLogFile.write("</tr>")
end if
Next
objLogFile.write("</table>")	

'--------------------------------------------------------------------------------------
    objLogFile.write("<p class=MsoNormal>" & "___________________________" & "</p>")
    objLogFile.write("<p class=MsoNormal><b>" & "Installation Report Complete for " & StrComputer & " at " & Date() & " " & Time() & "</b></p>")

    objLogFile.write "</div>"
    objLogFile.write "</body>"
    objLogFile.write "</html>"

     'wscript.echo "Done"
