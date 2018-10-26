Attribute VB_Name = "DNLabelsModule"

Public Function DNLabels() As Boolean
    Dim uName As String
    Dim uPass As String
    Dim RoutePartitions As Variant
    
    uName = UserForm1.ADName.text
    uPass = UserForm1.ADPassword.text
    
    If PublicFunctions.CheckUserPass(uName, uPass) = False Then
        DNLabels = False
        Exit Function
    End If
    
    If PublicFunctions.VerifyConnectivity(uName, uPass) = False Then
        MsgBox ("Unable to Verify Connectivity." & vbCrLf & "Please check username and password." _
            & vbCrLf & "If issue persist verify connectivity to CUCM.")
        DNLabels = False
        Exit Function
    End If
    
    RoutePartitions = listRoutePartition(uName, uPass)
    If IsEmpty(RoutePartitions) Then
        MsgBox ("Unable to retrieve list of route partitions." & vbCrLf & "Check CUCM privlidges")
        FullReport = False
        Exit Function
    End If
    
    
    
    
    
End Function


Private Function listRoutePartition(uName As String, uPass As String) As Variant
 'Set and instantiate our working objects
    Dim Req As Object
    Dim sEnv As String
    Dim Results() As String
    Dim RoutePartitionsArr As Variant
    Set Req = CreateObject("MSXML2.ServerXMLHTTP")
    Req.Open "Post", "https://10.200.232.111:8443/axl", False, uName, uPass
    Req.setRequestHeader "cache-control", "no-cache"
    Req.setRequestHeader "content-type", "text/xml"
    Req.setRequestHeader "soapaction", "CUCM:DB ver=10.5 getPhone"
    
 ' we create our SOAP envelope for submission to the Web Service
     sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:ns=""http://www.cisco.com/AXL/API/10.5"">" & vbCrLf
     sEnv = sEnv & "    <soapenv:Header/>" & vbCrLf
     sEnv = sEnv & "    <soapenv:Body>" & vbCrLf
     sEnv = sEnv & "        <ns:listRoutePartition>" & vbCrLf
     sEnv = sEnv & "            <searchCriteria>" & vbCrLf
     sEnv = sEnv & "                <name>%</name>" & vbCrLf
     sEnv = sEnv & "            </searchCriteria>" & vbCrLf
     sEnv = sEnv & "            <returnedTags>" & vbCrLf
     sEnv = sEnv & "                <name></name>" & vbCrLf
     sEnv = sEnv & "            </returnedTags>" & vbCrLf
     sEnv = sEnv & "        </ns:listRoutePartition>" & vbCrLf
     sEnv = sEnv & "    </soapenv:Body>" & vbCrLf
     sEnv = sEnv & "</soapenv:Envelope>" & vbCrLf
' Send SOAP Request
    'MsgBox sEnv
    Req.send sEnv
    
' Output Results
    RoutePartitions() = Split(PublicFunctions.RegexExtract(Req.responseText, "name"), ",")
    If UBound(RoutePartitions) = -1 Then
        Exit Function
    End If
    Dim i As Integer
    ReDim RoutePartitionsArr(LBound(RoutePartitions()) To UBound(RoutePartitions()))
    For i = LBound(RoutePartitions()) To UBound(RoutePartitions())
        RoutePartitionsArr(i) = RoutePartitions(i)
    Next i
    
  'clean up code
    Set Req = Nothing
    listRoutePartition = RoutePartitionsArr
    
End Function

