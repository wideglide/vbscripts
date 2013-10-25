'==========================================================================================
'  Write to an Excel file
'==========================================================================================
Dim objXLApp, objXLWb, objXLWs

strFile = "test.xlsx"
Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory
Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = True
Set objXLWb = objXLApp.Workbooks.Open(strPath & "\" & strFile)

'~~> Working with Sheet1
Set objXLWs = objXLWb.Sheets(1)

intRow = 1
Do Until objXLWs.Cells(intRow,1).Value = ""
  Wscript.Echo objXLWs.Cells(intRow,1)
  Call UpdateInfo(2,intRow)
  intRow = intRow + 1
Loop
  

objXLWb.Save
'~~> Save as Excel File (xls) to retain format
'objXLWb.SaveAs "C:\Users\joshua.bundt\Downloads\scripts\test1.xlsx", 51

'~~> File Formats
'51 = xlOpenXMLWorkbook (without macro's in 2007-2010, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2010, xlsm)
'50 = xlExcel12 (Excel Binary Workbook in 2007-2010 with or without macro's, xlsb)
'56 = xlExcel8 (97-2003 format in Excel 2007-2010, xls)

objXLWb.Close(SaveChanges=False)

Set objXLWs = Nothing
Set objXLWb = Nothing

objXLApp.Quit
Set objXLApp = Nothing


Sub CheckBitlocker
  strComputer = "." 
  Set objShell = CreateObject("Wscript.Shell") 
  strEnvSysDrive = objShell.ExpandEnvironmentStrings("%SystemDrive%") 
   
  Set objWMIServiceBit = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption") 
  Set colItems = objWMIServiceBit.ExecQuery("SELECT * FROM Win32_EncryptableVolume",,48) 
   
  For Each objItem in colItems 
      If objItem.DriveLetter = strEnvSysDrive Then 
          strDeviceC = objItem.DeviceID 
          DriveC =  "Win32_EncryptableVolume.DeviceID='"&strDeviceC&"'" 
          Set objOutParams = objWMIServiceBit.ExecMethod(DriveC, "GetProtectionStatus") 
          If objOutParams.ProtectionStatus = "1" Then 
              wscript.Echo "Bitlocker is enabled" 
          Else 
              wscript.Echo "Bitlocker is disabled" 
          End if 
      End If 
  Next
End Sub

Sub CheckBitlocker2
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption")
  Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_EncryptableVolume",,48)
  For Each objItem in colItems
  	Wscript.Echo "-----------------------------------"
  	Wscript.Echo "Win32_EncryptableVolume instance"
  	Wscript.Echo "-----------------------------------"
  	Wscript.Echo "ProtectionStatus: " & objItem.ProtectionStatus
  Next
End Sub

Sub UpdateInfo(intSheet, intRow)
  strComputer = objXLWs.Cells(intRow,1)
  'Write data to sheet 2
  Set objWs = objXLWb.Sheets(intSheet)
  Set Network = CreateObject("WScript.Network")
  Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
  For Each item In query_result
   objWs.Cells(intRow,2) = item.Vendor
   objWs.Cells(intRow,3) = item.Name
   objWs.Cells(intRow,1) = item.IdentifyingNumber
   objWs.Cells(intRow,6) = item.UUID
  Next
  objWs.Cells(intRow,4) = Network.ComputerName
  objWs.Cells(intRow,5) = Network.Username
  set query_result = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive",,48)
  driveModel = ""
  driveSerial = ""
  driveSize = ""
  For Each item In query_result
    If driveModel <> "" Then comma = ", " Else comma = "" End If
    driveModel = driveModel & comma & Trim(item.Model)
    driveSerial = driveSerial & comma & Trim(item.SerialNumber)
    driveSize = driveSize & comma & Trim(item.Size)
  Next
  netMAC = ""
  set query_result = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True",,48)
  For Each item in query_result
    if netMac <> "" Then comma = ", " Else comma = "" End If
    netMAC = netMAC & comma & Trim(item.MACAddress)
  Next
  objWs.Cells(intRow,7) = driveModel
  objWs.Cells(intRow,8) = driveSerial
  objWs.Cells(intRow,9) = driveSize
  objWs.Cells(intRow,10) = netMAC
  
  strQuery = "winmgmts:{impersonationLevel=impersonate,AuthenticationLevel=pktprivacy}\\" & strComputer & "\root\CIMV2\Security\MicrosoftTpm"
  Set objWMI = GetObject(strQuery)
  Set query_result = objWMI.InstancesOf("Win32_Tpm")
  For Each item In query_result
    If item.IsEnabled(A) Then objWs.Cells(intRow,11) = A End If
    If item.IsActivated(B) Then objWs.Cells(intRow,12) = B End If
    If item.IsOwned(C) Then objWs.Cells(intRow,13) = C End If
    objWs.Cells(intRow,14) = item.SpecVersion
  Next
  
  strQuery = "winmgmts:\\" & strComputer & "\root\CIMV2\Security\MicrosoftVolumeEncryption"
  Set objWMI = GetObject(strQuery)
  Set query_result = objWMI.InstancesOf("Win32_EncryptableVolume")
  For Each item in query_result
    objWs.Cells(intRow,15) = item.ProtectionStatus
  Next
  Set query_result = Nothing
  Set objWMI = Nothing
  Set Network = Nothing
  Set objWs = Nothing
End Sub