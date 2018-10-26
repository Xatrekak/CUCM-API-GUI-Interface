Attribute VB_Name = "GetPhoneModule"
Public Function GetPhone() As Boolean

    Dim uName As String
    Dim uPass As String
    Dim IPQuery As Boolean
    Dim PhoneIPorMAC As String
    Dim PhoneIPorMAC_property As String
    Dim PhoneNames As Variant
    Dim PhoneAttrb1 As Variant
    Dim PhoneAttrb2 As Variant
    ReDim PhoneAttrb1(0 To 4) As Variant
    ReDim PhoneAttrb2(0 To 4) As Variant
    ReDim InputCheck(0 To 2) As Variant
    
    uName = UserForm1.ADName.text
    uPass = UserForm1.ADPassword.text
    
    If PublicFunctions.CheckUserPass(uName, uPass) = False Then
        GetPhone = False
        Exit Function
    End If
    
    If PublicFunctions.VerifyConnectivity(uName, uPass) = False Then
        MsgBox ("Unable to Verify Connectivity." & vbCrLf & "Please check username and password." _
            & vbCrLf & "If issue persist verify connectivity to CUCM.")
        GetPhone = False
        Exit Function
    End If
    
    PhoneIPorMAC = UserForm1.PhoneIPorMAC.text
    If PhoneIPorMAC = "" Then
        MsgBox ("You did not enter an IP or MAC Address.")
        GetPhone = False
        Exit Function
    End If
    
    PhoneIPorMAC_property = PublicFunctions.IPorMACAddrValidator(PhoneIPorMAC)
    If PhoneIPorMAC_property = "MACAddr" Then
     PhoneIPorMAC = "SEP" & UCase(PhoneIPorMAC)
     IPQuery = False
    ElseIf PhoneIPorMAC_property = "IPAddr" Then
        IPQuery = True
    Else
    MsgBox ("It seems that you entereed something besides an IP or MAC Address." _
            & vbCrLf & "Please check your input and try again.")
    GetPhone = False
    Exit Function
    End If
    
    ThisWorkbook.Sheets("Template").Copy Before:=Workbooks.Add.Sheets(1)
    ActiveSheet.Name = "Report_Output"
    Sheets("Sheet1").Delete
    Dim i As Integer
    Dim j As Integer
    j = 2
    PhoneAttrb2 = PublicFunctions.RisPortQuery(CStr(PhoneIPorMAC), uName, uPass, IPQuery)
    If PhoneAttrb2(0) <> "" Then
        Range(Cells(j, 1), Cells(j, 5)) = PhoneAttrb2
        PhoneAttrb1 = PublicFunctions.listPhone(CStr(PhoneAttrb2(0)), uName, uPass)
        Range(Cells(j, 6), Cells(j, 10)) = PhoneAttrb1
        j = j + 1
    End If
    DoEvents
        
    GetPhone = True
    
End Function
