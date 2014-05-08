 on error resume next
'
' copyright 2012
' all rights reserved
'
'
Const HKEY_LOCAL_MACHINE = &H80000002
Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
dim x, strComp, strComputer, objRecordSet

'
' adjust strBook
'
 strPath = inputbox("Type or Paste path to file:")
  strFile = inputbox("Enter Excel file name: (include extention)")
  strBook = strPath &"\"& strFile

Set gExcel = CreateObject("Excel.Application")
gExcel.visible = true
gExcel.Workbooks.Open(strBook)

set objShell = Wscript.CreateObject("Wscript.Shell")
Set objUserInfo = CreateObject("Scripting.Dictionary")
Set objAdsDict = CreateObject("Scripting.Dictionary")
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objRank = CreateObject("Scripting.Dictionary")
objDictionary.CompareMode = TextMode
objAdsDict.CompareMode = TextMode
objUserInfo.CompareMode = TextMode
objRank.CompareMode = TextMode

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = _
    "SELECT cn, info, AdsPath FROM 'LDAP://ou=BDE,ou=TF AVIATION,ou=Computers - STIG,ou=HD_DSST,ou=KDHR,ou=RC South,dc=afghan,dc=swa,dc=ds,dc=army,dc=mil' WHERE objectCategory='computer'"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	strName = objRecordSet.Fields("cn").value
	strModel = formatInfo(objRecordSet.Fields("info").value)
	strAdsPath = objRecordSet.Fields("AdsPath").value
	objDictionary.Add strName, strModel
	objAdsDict.Add strName, strAdsPath
	objRecordSet.MoveNext
Loop
objRecordSet = Nothing
'
'Move through dictionary object and compare to current spreadsheet
'
intRowCount = gExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
wscript.echo "first intRowCount " & intRowCount

for x = 2 to intRowCount
 strComp = gExcel.Sheets(1).Cells(x,1).value
   if objDictionary.Exists(strComp) then
     gExcel.Sheets(1).Cells(x,8).value = updateSerial(objDictionary.Item(strComp))
     gExcel.Sheets(1).Cells(x,11).value = updateMake(objDictionary.Item(strComp))
     gExcel.Sheets(1).Cells(x,12).value = updateModel(objDictionary.Item(strComp))
     gExcel.Sheets(1).Cells(x,10).value = updateLLogon(objAdsDict.Item(strComp))
     gExcel.Sheets(1).Cells(x,17).value = "YES"
     objDictionary.Remove(strComp)
   else
     gExcel.Sheets(1).Cells(x,17).value = "NO"
   end if
   if (DateDiff("d", gExcel.Sheets(1).Cells(x,13).value, Now())) > 13 or _
	gExcel.Sheets(1).Cells(x,15).value = "not found" then
     gExcel.Sheets(1).Cells(x,14).value = "stale"
   end if
next

y=intRowCount +1

colKeys = objDictionary.Keys
for Each strKey in colKeys
 gExcel.Sheets(1).Cells(y,1).value = strKey
 gExcel.Sheets(1).Cells(y,8).value = updateSerial(objDictionary.Item(strKey))
 gExcel.Sheets(1).Cells(y,11).value = updateMake(objDictionary.Item(strKey))
 gExcel.Sheets(1).Cells(y,12).value = updateModel(objDictionary.Item(strKey))
 gExcel.Sheets(1).Cells(y,10).value = updateLLogon(objAdsDict.Item(strKey))

 gExcel.Sheets(1).Cells(y,13).value = Now()
 gExcel.Sheets(1).Cells(y,14).value = "new"
 y=y+1
next

'
'
'now for the meat of the scrip, we have all the current computer name
'and have identified stale or new ones.  Time to make the donuts
'
'
intRowCount = gExcel.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
wscript.echo "after intRowCount " & intRowCount
for x = 2 to intRowCount
 strBubba = gExcel.Sheets(1).Cells(x,1).value
 if gExcel.Sheets(1).Cells(x,14).value <> "updated" then
  if checkPing(strBubba) = "good" then 
  gExcel.Sheets(1).Cells(x,15).value = "pung"
  gExcel.Sheets(1).Cells(x,3).value = updateLUser(strBubba)
  gExcel.Sheets(1).Cells(x,4).value = updateUser(strBubba)
  gExcel.Sheets(1).Cells(x,7).value = updateMac(strBubba)
  else
  gExcel.Sheets(1).Cells(x,14).value = "stale"
  gExcel.Sheets(1).Cells(x,15).value = "not found"
  end if
'
'check for errors
'
  if gExcel.Sheets(1).Cells(x,3).value = "error" or gExcel.Sheets(1).Cells(x,3).value = "" then
   gExcel.Sheets(1).Cells(x,14).value = "error"
  elseif gExcel.Sheets(1).Cells(x,4).value = "error" or gExcel.Sheets(1).Cells(x,4).value = "" then
   gExcel.Sheets(1).Cells(x,14).value = "error"
  elseif gExcel.Sheets(1).Cells(x,7).value = "error" or gExcel.Sheets(1).Cells(x,7).value = "" then
   gExcel.Sheets(1).Cells(x,14).value = "error"
  elseif gExcel.Sheets(1).Cells(x,10).value = "error" or gExcel.Sheets(1).Cells(x,10).value = "" then
   gExcel.Sheets(1).Cells(x,14).value = "error"

  elseif gExcel.Sheets(1).Cells(x,15).value = "not found" or gExcel.Sheets(1).Cells(x,15).value = "" then
   gExcel.Sheets(1).Cells(x,14).value = "error"

  else
   gExcel.Sheets(1).Cells(x,13).value = Now()
   gExcel.Sheets(1).Cells(x,14).value = "updated"
  end if

 else ' checkping else

  gExcel.Sheets(1).Cells(x,14).value = "not updated " & Now()

 end if 'checkping end
next
'
'
' use the user info to populate people
'
'
objCommand.CommandText = _
    "SELECT sn, givenName, initials, mail, sAMAccountName, title FROM 'LDAP://ou=BDE,ou=TF AVIATION,ou=USERS,ou=HD_DSST,ou=KDHR,ou=RC South,dc=afghan,dc=swa,dc=ds,dc=army,dc=mil' WHERE objectCategory='user'"  
Set objRecordSet = objCommand.Execute


objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    strUserKey = "NANW\" & objRecordSet.Fields("sAMAccountName").Value
    strUserData = objRecordSet.Fields("sn").Value &", "& _
      objRecordSet.Fields("givenName").Value &" "& _
      objRecordSet.Fields("initials").Value &":"& _
      objRecordSet.Fields("mail").Value
     strRank = objRecordSet.Fields("title").Value
      objUserInfo.add strUserKey,strUserData
      objRank.add strUserKey,strRank
    objRecordSet.MoveNext
Loop

for x = 2 to intRowCount
 strUserInfo1 = gExcel.Sheets(1).Cells(x,3).value
 if objUserInfo.Exists(strUserInfo1) then
     gExcel.Sheets(1).Cells(x,5).value = Mid(objUserInfo.Item(strUserInfo1),1,(InStr(objUserInfo.Item(strUserInfo1),":")-1))
     gExcel.Sheets(1).Cells(x,6).value = Mid(objUserInfo.Item(strUserInfo1),(InStr(objUserInfo.Item(strUserInfo1),":")+1),(Len(objUserInfo.Item(strUserInfo1))-(InStr(objUserInfo.Item(strUserInfo1),":"))))
     gExcel.Sheets(1).Cells(x,16).value = objRank.Item(strUserInfo1)
  else 
     gExcel.Sheets(1).Cells(x,5).value = "not in 7ID OU"
     gExcel.Sheets(1).Cells(x,6).value = "not in 7ID OU"
     gExcel.Sheets(1).Cells(x,16).value = "not in 7ID OU"
 end if
next

gExcel.save
gExcel.close
gExcel.quit
gExcel = nothing

Wscript.Echo "Done"
'
'subroutines and functions
'
function updateUser(ByVal strComputer)
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
if Err.Number <> 0 then
 tmpErr = Err.Number & ":" & Err.Description
 wscript.echo "error happened in updateUser for: " & strComputer & ":"& tmpErr
 updateUser = "error"
 Err.Clear
Set objWMIService = nothing
 Exit Function
else
Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
 For Each objComputer in colComputer
  strUserTmp = CStr(objComputer.UserName)
 Next
 updateUser = strUserTmp
end if
Set objWMIService = nothing
end function

function updateLUser(ByVal strComputer)
 Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
  if Err.Number <> 0 then
   tmpErr = Err.Number & ":" & Err.Description
   wscript.echo "error happened in updateLUser for: " & strComputer & ":"& tmpErr
   updateLUser = "error"
   Err.Clear
Set objWMIService = nothing
   Exit Function
  else
   strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
   strValueName = "LastLoggedOnSAMUser"
   objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
   strLUserTemp = strValue
   updateLUser = strLUserTemp
  end if
Set objWMIService = nothing
end function

function updateMac(ByVal tmpMac)
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & tmpMac & "\root\cimv2")
if Err.Number <> 0 then
 tmpErr = Err.Number & ":" & Err.Description
 wscript.echo "error happened in updateMac for: " & tmpMac & ":"& tmpErr
 updateMac = "error"
 Err.Clear
Set objWMIService = nothing
 Exit Function
else
 Set colAdapters = objWMIService.ExecQuery _
     ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
  For Each objAdapter in colAdapters
    strMac = CStr(objAdapter.MACAddress)
  Next
 updateMac = strMac
end if
Set objWMIService = nothing
end function

function updateSerial(tmp0)
 tmpLength = Len(tmp0)
 tmp1 = Mid(tmp0, InStr(tmp0,":")+1, tmpLength-InStr(tmp0,":")-1)
 updateSerial = tmp1
end function

function updateLLogon(comp0)
  Set objComp = GetObject(comp0)
  if Err.Number <> 0 then
   updateLLogon = "error"
   Err.Clear
   Exit Function
  else
   Set objLastLogon = objComp.Get("lastLogonTimestamp")
   intLastLogonTime = objLastLogon.HighPart * (2^32) + objLastLogon.LowPart 
   intLastLogonTime = intLastLogonTime / (60 * 10000000)
   intLastLogonTime = intLastLogonTime / 1440
   updateLLogon = (intLastLogonTime + #1/1/1601#)
  end if
end function

function updateMake(tmp0)
 tmp1 = Left(tmp0,InStr(tmp0,"|")-1)
 updateMake = tmp1
end function

function updateModel(tmp0)
 tmp2 = Mid(tmp0, InStr(tmp0,"|")+1,InStr(tmp0,":")-(InStr(tmp0,"|"))-1)
 updateModel = tmp2
end function

Function checkPing(comp)
Set objExecObject = objShell.Exec("cmd /c ping -n 2 -w 1000 " & comp)
Do while Not objExecObject.StdOut.AtEndOfStream
	strText = objExecObject.StdOut.ReadLine()
       If (Instr(strText, "Reply") > 0 and Instr(strText, "unreachable") = 0) Then
	checkPing = "good"
       End If
loop
End Function
'
' information
' InStr([start,]string1,string2[,compare])  compare 0 binary, 1 text
'

Function formatInfo(str)
 If str <> "" then
  make = Mid(str, 5, (InStr(str,"|"))-5)  
  model = Mid(str, (InStr(str,"|"))+1, (InStr(str,";"))-((InStr(str,"|"))+1))
  serial = Mid(str, (InStr(str,"SN="))+3,20 )
  serial1 = Mid(serial, 1, (InStr(serial,";")))
 formatInfo = make & "|" & model & ":" & serial1
 else
  formatInfo = "unknown|unknown:unknown"
 end if
end Function
