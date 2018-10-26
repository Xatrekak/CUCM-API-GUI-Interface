VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15570
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    With Me
       .Width = 750
       .Height = 450
    End With
End Sub



Private Sub FullReportButton_Click()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
    If MsgBox("Press Ok to launch the Full Report." & vbCrLf & "This report will take an extended period of time" _
            & vbCrLf & "You will get an alert once it has finished running.", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    
    If FullReportModule.FullReport = False Then Exit Sub
    
    MsgBox ("Report has finsihed")
End Sub


Private Sub GetPhoneButton_Click()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
    If MsgBox("You will get an alert once it has finished running.", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    
    If GetPhoneModule.GetPhone = False Then Exit Sub
    
    MsgBox ("Report has finsihed")
End Sub



Private Sub UserForm_Terminate()
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub

