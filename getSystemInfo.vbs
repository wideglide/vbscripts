'Option Explicit

'Get the current user service tag
'strComputer = "LEWINB16AV1EB01"
strComputer = "."

Const wbemFlagReturnImmediately = 16
Const wbemFlagForwardOnly = 32
lFlags = wbemFlagReturnImmediately + wbemFlagForwardOnly

Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_BIOS")
For Each item In query_result
 strServiceTag = item.SerialNumber
 wscript.echo "Win32_BIOS.SerialN: " & strServiceTag
 Wscript.echo "Win32_BIOS.Manufac: " & item.Manufacturer
 Wscript.echo "Win32_BIOS.Name   : " & item.Name
Next

Set query_result = objWMI.ExecQuery("SELEcT * from Win32_SystemEnclosure")
For Each item In query_result
 Wscript.echo "*** Win32_SystemEnclosure"
 Wscript.echo "Part Number    : " & item.PartNumber
 Wscript.echo "Serial Number  : " & item.SerialNumber
 Wscript.echo "Asset Tag      : " & item.SMBIOSAssetTag
 Wscript.echo "Manufacturer   : " & item.Manufacturer
 Wscript.echo "Model          : " & item.Model
 Wscript.echo "Name           : " & item.Name
 Wscript.echo "SKU            : " & item.SKU
Next

Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each item In query_result
 Wscript.echo "*** Win32_ComputerSystem"
 Wscript.echo "Manufacturer   : " & item.Manufacturer
 Wscript.echo "Model          : " & item.Model
 Wscript.echo "Name           : " & item.Name
Next

Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
For Each item In query_result
 Wscript.echo "*** Win32_ComputerSystemProduct"
 Wscript.echo "Identifying Num: " & item.IdentifyingNumber
 Wscript.echo "SKU Number     : " & item.SKUNumber
 Wscript.echo "Name           : " & item.Name
 Wscript.echo "UUID           : " & item.UUID
 Wscript.echo "Vendor         : " & item.Vendor
Next

Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive",,lFlags)
For Each item In query_result
 Wscript.echo "*** Win32_DiskDrive"
 Wscript.echo "Model          : " & item.Model
 Wscript.echo "SerialNumber   : " & item.SerialNumber
 Wscript.echo "Size           : " & item.Size
Next

Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each item In query_result
  Wscript.Echo "***Win32_NetworkAdapterConfiguration"
  Wscript.Echo "MAC Address   : " & item.MACAddress
Next

Set query_result = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetEnabled = True")
For Each item In query_result
  Wscript.Echo "***Win32_NetworkAdapter"
  Wscript.Echo "MAC Address   : " & item.MACAddress
  Wscript.Echo "ProductName   : " & item.ProductName
Next

'Get the Product ID from WMI
strService = "winmgmts:{impersonationlevel=impersonate}//./root/HP/InstrumentedBIOS"
strQuery = "select * from HP_BIOSSetting where Name='Product Number' or Name='SKU Number'"
Set objWMI = GetObject(strService)
Set colItems = objWMI.ExecQuery(strQuery,,lFlags)
sProductNumber = ""
For Each objItem In colItems
    sProductNumber = objItem.Value
    Wscript.Echo objItem.Name & " : " & objItem.Value
Next

set objWMI = Nothing


'Call EnumNameSpaces("root")
'Call EnumClass(".","\root\HP\InstrumentedBIOS")
'Call EnumClasses("\root\HP\InstrumentedBIOS")
'Call EnumClassProperties(".","\root\HP\InstrumentedBIOS","HP_BIOSSetting")
'Call EnumClassProperties(".","\root\cimv2","Win32_BIOS")

Sub EnumNameSpaces(strNameSpace)
    On Error Resume Next
    WScript.Echo strNameSpace
    Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\" & strNameSpace)

    Set colNameSpaces = objWMIService.InstancesOf("__NAMESPACE")

    For Each objNameSpace in colNameSpaces
        Call EnumNameSpaces(strNameSpace & "\" & objNameSpace.Name)
    Next
End Sub

Sub EnumClass(strComputer,strNameSpace)
 Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & strNameSpace)

 For Each objclass in objWMIService.SubclassesOf()
    'WScript.Echo objclass.Path_.Class
     Call EnumClassProperties(strComputer, strNameSpace, objclass.Path_.Class)
 Next
End Sub

Sub EnumClasses(strNameSpace)

 strComputer = "."
 Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & strNameSpace)

 For Each objclass in objWMIService.SubclassesOf()
    WScript.StdOut.Write objclass.Path_.Class
    arrDerivativeClasses = objClass.Derivation_ 
    For Each strDerivativeClass in arrDerivativeClasses 
       WScript.StdOut.Write " <- " & strDerivativeClass
    Next
    WScript.StdOut.Write vbNewLine
 Next

    For Each objNameSpace in colNameSpaces
        Call EnumNameSpaces(strNameSpace & "\" & objNameSpace.Name)
    Next
End Sub

Sub EnumClassProperties(strComputer, strNameSpace, strClass)
 Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & strNameSpace & ":" & strClass)

 WScript.Echo strClass & " Class Properties"
 WScript.Echo "------------------------------"

 For Each objClassProperty in objClass.Properties_
    WScript.Echo objClassProperty.Name
 Next

End Sub

