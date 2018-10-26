Attribute VB_Name = "PublicFunctions"
Public Function IPorMACAddrValidator(InputToCheck) As String

    Dim allMatches As Object
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    Dim i As Long
    Dim result As String
    Dim RegPattern As Variant
    ReDim RegPattern(0 To 1) As Variant
    

    RegPattern(0) = "^(([0-9A-Fa-f]{2}([\.:-]?)){5}([0-9A-Fa-f]{2}))$"
    RegPattern(1) = "^((?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))$"
    For Each j In RegPattern
    
        RE.Pattern = j
        RE.Global = True
        Set allMatches = RE.Execute(InputToCheck)
        
        
        For i = 0 To allMatches.Count - 1
            For k = 0 To allMatches.Item(i).submatches.Count - 1
                result = allMatches.Item(i).submatches.Item(k)
            Next
        Next
        
        If Len(result) <> 0 Then
            If j = RegPattern(0) Then
                IPorMACAddrValidator = "MACAddr"
                Exit Function
            Else
                IPorMACAddrValidator = "IPAddr"
                Exit Function
            End If
        End If
    Next
    IPorMACAddrValidator = ""
    Set RE = Nothing

End Function


Public Function CheckUserPass(uName As String, uPass As String) As Boolean
    Dim MissingInput As String
    
    If uName = "" Then: MissingInput = "AD Name"
    If uPass = "" Then
        If MissingInput = "" Then
            MissingInput = "AD Password"
        Else
            MissingInput = MissingInput & " And Password"
        End If
    End If
    
    
    If MissingInput <> "" Then
    CheckUserPass = False
    MsgBox ("Check Your inputs: " & MissingInput & " appears to be empty.")
    Exit Function
    End If
    
    CheckUserPass = True
    
End Function

Public Function RisPortQuery(phoneName As String, uName As String, uPass As String, Optional IPQuery As Boolean) As Variant
    Dim Req As Object
    Dim sEnv As String
    Dim PhoneAttrb2 As Variant
    Dim tempArray As Variant
    ReDim PhoneAttrb2(0 To 4) As Variant
    ReDim tempArray(0 To 4) As Variant
    Dim StoredResponse As String
    PhoneAttrb2 = Array("Name", "IpAddress", "DirNumber", "Status", "StatusReason")
    
    If IPQuery = True Then SelectByValue = "IpAddress" Else SelectByValue = "Name"
    
    Set Req = CreateObject("MSXML2.ServerXMLHTTP")
    Req.Open "Post", "https://10.200.232.111:8443/realtimeservice/services/RisPort", False, uName, uPass
    Req.setRequestHeader "cache-control", "no-cache"
    Req.setRequestHeader "content-type", "text/xml"
    Req.setRequestHeader "soapaction", "CUCM:DB ver=10.5 getPhone"
    
    sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:soapenc=""http://schemas.xmlsoap.org/soap/encoding/"">" & vbCrLf
    sEnv = sEnv & " <soapenv:Body>" & vbCrLf
    sEnv = sEnv & "     <SelectCmDevice >" & vbCrLf
    sEnv = sEnv & "         <CmSelectionCriteria>" & vbCrLf
    sEnv = sEnv & "             <MaxReturnedDevices>100000</MaxReturnedDevices>" & vbCrLf
    sEnv = sEnv & "             <Class>Any</Class>" & vbCrLf
    sEnv = sEnv & "             <Model>255</Model>" & vbCrLf
    sEnv = sEnv & "             <SelectBy>" & SelectByValue & "</SelectBy>" & vbCrLf
    sEnv = sEnv & "                 <SelectItems soapenc:arrayType=""SelectItem[66]"" >" & vbCrLf
    sEnv = sEnv & "                     <item><Item>" & phoneName & "</Item></item>" & vbCrLf
    sEnv = sEnv & "                 </SelectItems>" & vbCrLf
    sEnv = sEnv & "         </CmSelectionCriteria>" & vbCrLf
    sEnv = sEnv & "     </SelectCmDevice>" & vbCrLf
    sEnv = sEnv & " </soapenv:Body>" & vbCrLf
    sEnv = sEnv & "</soapenv:Envelope>"
        
    Req.send sEnv
    StoredResponse = Req.responseText
    For i = LBound(PhoneAttrb2) To UBound(PhoneAttrb2)
        tempArray(i) = RegexExtract(StoredResponse, CStr(PhoneAttrb2(i)), True)
    Next i
    Set Req = Nothing
    RisPortQuery = tempArray

End Function


Public Function listPhone(phoneName As String, uName As String, uPass As String) As Variant
    Dim Req As Object
    Dim sEnv As String
    Dim PhoneAttrb1 As Variant
    Dim tempArray As Variant
    ReDim PhoneAttrb1(0 To 4) As Variant
    ReDim tempArray(0 To 4) As Variant
    Dim StoredResponse As String
    PhoneAttrb1 = Array("description", "model", "protocol", "disableSpeaker", "voiceVlanAccess")
    
    
    Set Req = CreateObject("MSXML2.ServerXMLHTTP")
    Req.Open "Post", "https://10.200.232.111:8443/axl", False, uName, uPass
    Req.setRequestHeader "cache-control", "no-cache"
    Req.setRequestHeader "content-type", "text/xml"
    Req.setRequestHeader "soapaction", "CUCM:DB ver=10.5 getPhone"
    
    sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns=""http://www.cisco.com/AXL/API/10.5"">" & vbCrLf
    sEnv = sEnv & " <soapenv:Header/>" & vbCrLf
    sEnv = sEnv & " <soapenv:Body>" & vbCrLf
    sEnv = sEnv & "     <ns:getPhone>" & vbCrLf
    sEnv = sEnv & "         <name>" & phoneName & "</name>" & vbCrLf
    sEnv = sEnv & "     </ns:getPhone>" & vbCrLf
    sEnv = sEnv & " </soapenv:Body>" & vbCrLf
    sEnv = sEnv & "</soapenv:Envelope>" & vbCrLf
    
    Req.send sEnv
    StoredResponse = Req.responseText
    For i = LBound(PhoneAttrb1) To UBound(PhoneAttrb1)
        tempArray(i) = RegexExtract(StoredResponse, CStr(PhoneAttrb1(i)))
        If PhoneAttrb1(i) = "disableSpeaker" Then
            If tempArray(i) = False Then
                tempArray(i) = "Enabled"
            Else: tempArray(i) = "Disabled"
            End If
        End If
        If PhoneAttrb1(i) = "voiceVlanAccess" Then
            If tempArray(i) = "0" Then
                tempArray(i) = "Enabled"
            Else: tempArray(i) = "Disabled"
            End If
        End If
    Next i
    Set Req = Nothing
    listPhone = tempArray
    
End Function


Public Function VerifyConnectivity(uName As String, uPass As String) As Boolean
    Dim Req As Object
    Dim AXLStatus(0 To 0) As Variant
    Dim StoredResponse As String
    
    
    Set Req = CreateObject("MSXML2.ServerXMLHTTP")
    Req.Open "Get", "https://10.200.232.111:8443/axl", False, uName, uPass
    Req.setRequestHeader "cache-control", "no-cache"
    Req.setRequestHeader "content-type", "text/xml"
    Req.setRequestHeader "soapaction", "CUCM:DB ver=10.5 getPhone"
    
    Req.send ""
    
    StoredResponse = Req.responseText
    
    AXLStatus(0) = RegexExtract(StoredResponse, "The (AXL Web Service is working) and accepting requests.", GeneralQuery:=True)
    If AXLStatus(0) = "" Then VerifyConnectivity = False Else VerifyConnectivity = True
    
    
End Function

'This function requires FieldToExtract to also contain a capture group. It will not extract pure text
Function RegexExtract(InputToExtract As String, FieldToExtract As String, Optional RisPortInput As Boolean, Optional GeneralQuery As Boolean)
Dim allMatches As Object
Dim RE As Object
Dim separator As String
Set RE = CreateObject("vbscript.regexp")
Dim i As Long, j As Long
Dim result As String
separator = ","


If GeneralQuery = True Then
    RE.Pattern = FieldToExtract
    Else
        If RisPortInput = False Then
            RE.Pattern = "<" & FieldToExtract & ">(.*?)</" & FieldToExtract & ">"
        ElseIf FieldToExtract = "Name" Then
            RE.Pattern = "<" & FieldToExtract & ".*?>(SEP.*?)</" & FieldToExtract & ">"
        Else
            RE.Pattern = "<" & FieldToExtract & " .*?>(.*?)</" & FieldToExtract & ">"
        End If
    End If


RE.Global = True
Set allMatches = RE.Execute(InputToExtract)
For i = 0 To allMatches.Count - 1
    For j = 0 To allMatches.Item(i).submatches.Count - 1
        result = result & (separator & allMatches.Item(i).submatches.Item(j))
    Next
Next

If Len(result) <> 0 Then
    result = Right$(result, Len(result) - Len(separator))
End If

RegexExtract = result
Set RE = Nothing
End Function
