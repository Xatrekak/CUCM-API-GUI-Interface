Attribute VB_Name = "FullReportModule"
Public Function FullReport() As Boolean
    Dim uName As String
    Dim uPass As String
    Dim PhoneNames As Variant
    Dim PhoneAttrb1 As Variant
    Dim PhoneAttrb2 As Variant
    ReDim PhoneAttrb1(0 To 4) As Variant
    ReDim PhoneAttrb2(0 To 4) As Variant
    ReDim InputCheck(0 To 2) As Variant
    
    uName = UserForm1.ADName.text
    uPass = UserForm1.ADPassword.text
    
    If PublicFunctions.CheckUserPass(uName, uPass) = False Then
        FullReport = False
        Exit Function
    End If
    
    If PublicFunctions.VerifyConnectivity(uName, uPass) = False Then
        MsgBox ("Unable to Verify Connectivity." & vbCrLf & "Please check username and password." _
            & vbCrLf & "If issue persist verify connectivity to CUCM.")
        FullReport = False
        Exit Function
    End If
    
    PhoneNames = GetPhoneNames(uName, uPass)
    If IsEmpty(PhoneNames) Then
        MsgBox ("Unable to retrieve phone list." & vbCrLf & "Check CUCM privlidges")
        FullReport = False
        Exit Function
    End If
    ThisWorkbook.Sheets("Template").Copy Before:=Workbooks.Add.Sheets(1)
    ActiveSheet.Name = "Report_Output"
    Sheets("Sheet1").Delete
    Dim i As Integer
    Dim j As Integer
    j = 2
    For i = LBound(PhoneNames) To UBound(PhoneNames)
        PhoneAttrb2 = PublicFunctions.RisPortQuery(CStr(PhoneNames(i)), uName, uPass)
        If PhoneAttrb2(0) <> "" Then
            Range(Cells(j, 1), Cells(j, 5)) = PhoneAttrb2
            PhoneAttrb1 = PublicFunctions.listPhone(CStr(PhoneNames(i)), uName, uPass)
            Range(Cells(j, 6), Cells(j, 10)) = PhoneAttrb1
            j = j + 1
        End If
        DoEvents
    Next i
        
    FullReport = True
    
End Function



Private Function GetPhoneNames(uName As String, uPass As String) As Variant
 'Set and instantiate our working objects
    Dim Req As Object
    Dim sEnv As String
    Dim PhoneNames() As String
    Dim NameArr As Variant
    Set Req = CreateObject("MSXML2.ServerXMLHTTP")
    Req.Open "Post", "https://10.200.232.111:8443/axl", False, uName, uPass
    Req.setRequestHeader "cache-control", "no-cache"
    Req.setRequestHeader "content-type", "text/xml"
    Req.setRequestHeader "soapaction", "CUCM:DB ver=10.5 getPhone"
    
 ' we create our SOAP envelope for submission to the Web Service
     sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns=""http://www.cisco.com/AXL/API/10.5"">" & vbCrLf
     sEnv = sEnv & "    <soapenv:Header/>" & vbCrLf
     sEnv = sEnv & "    <soapenv:Body>" & vbCrLf
     sEnv = sEnv & "        <ns:listPhone>" & vbCrLf
     sEnv = sEnv & "            <searchCriteria>" & vbCrLf
     sEnv = sEnv & "                <name>SEP%</name>" & vbCrLf
     sEnv = sEnv & "            </searchCriteria>" & vbCrLf
     sEnv = sEnv & "            <returnedTags>" & vbCrLf
     sEnv = sEnv & "                <name></name>" & vbCrLf
     sEnv = sEnv & "            </returnedTags>" & vbCrLf
     sEnv = sEnv & "        </ns:listPhone>" & vbCrLf
     sEnv = sEnv & "    </soapenv:Body>" & vbCrLf
     sEnv = sEnv & "</soapenv:Envelope>" & vbCrLf
' Send SOAP Request
    'MsgBox sEnv
    Req.send sEnv
    
' Output Results
    PhoneNames() = Split(PublicFunctions.RegexExtract(Req.responseText, "name"), ",")
    If UBound(PhoneNames) = -1 Then
        Exit Function
    End If
    Dim i As Integer
    ReDim NameArr(LBound(PhoneNames()) To UBound(PhoneNames()))
    For i = LBound(PhoneNames()) To UBound(PhoneNames())
        NameArr(i) = PhoneNames(i)
    Next i
    
  'clean up code
    Set Req = Nothing
    GetPhoneNames = NameArr
    
End Function

