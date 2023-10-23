VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "Опции"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   OleObjectBlob   =   "UserForm11.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
 UserForm11.Hide
 UserForm4.Show
End Sub



Private Sub CommandButton12_Click()
    UserForm11.Hide
    UserForm6.Show
End Sub

Private Sub CommandButton14_Click()
    UserForm11.Hide
    DenisRequest
End Sub

Private Sub CommandButton15_Click()
    UserForm11.Hide
    Обновить
End Sub

Private Sub CommandButton16_Click()
    UserForm11.Hide
    LabourIntensity_SP_Query
End Sub

Private Sub CommandButton17_Click()
    UserForm11.Hide
    Components_SP_Query_
End Sub

Private Sub CommandButton18_Click()
    UserForm11.Hide
    UserForm7.Show
End Sub

Private Sub CommandButton19_Click()
    UserForm11.Hide
    UserForm8.Show
End Sub

Private Sub CommandButton20_Click()
    UserForm11.Hide
    UserForm9.Show
End Sub

Private Sub CommandButton21_Click()
    UserForm11.Hide
    UserForm10.Show
End Sub

Private Sub CommandButton22_Click()
    UserForm11.Hide
    Message
End Sub

Private Sub CommandButton23_Click()
    UserForm11.Hide
    UserForm7.Show
End Sub



Private Sub CommandButton4_Click()
    UserForm11.Hide
    UserForm5.Show
End Sub

Private Sub CommandButton5_Click()
    UserForm11.Hide
    UserForm3.Show
End Sub

Private Sub CommandButton27_Click()
    UserForm11.Hide
    GetProjectRefresh_SP_Query
End Sub

Private Sub CommandButton39_Click()
    UserForm11.Hide
    Aligment4d
End Sub

Private Sub CommandButton40_Click()
    UserForm11.Hide
    SaveToEXL
End Sub

Private Sub CommandButton6_Click()
    UserForm11.Hide
    TimeSheet.TimeSheet
End Sub

Private Sub Image1_Click()
 UserForm11.Hide
 SaveToPDF
End Sub


Private Sub Image11_Click()
    UserForm11.Hide
    aligment.aligment
End Sub

Private Sub Image12_Click()
    UserForm11.Hide
    Negotiation
End Sub

Private Sub Image13_Click()
    UserForm11.Hide
    Python
End Sub

Private Sub Image14_Click()
    UserForm11.Hide
    HideSys
End Sub


Private Sub Image15_Click()
    UserForm11.Hide
    ActiveWorkbook.Sheets("Задействование").Visible = True
End Sub


'
'Private Sub Image15_Click()
'    UserForm11.Hide
'    Me.CheckBox_Click
'End Sub

Private Sub Image17_Click()
    UserForm11.Hide
    UnhideSys
End Sub

Private Sub Image19_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image19_Click()
On Error Resume Next
    UserForm11.Hide
'    LabourIntensity_SP_Query
'    Components_SP_Query_
    Обновить
    aligment.aligment
    Обновить
    Aligment4d
    SaveToEXL
'    SaveToPDF
'    Python
End Sub

Private Sub Image6_Click()
    UserForm11.Hide
    Clone9
End Sub

Private Sub Image7_Click()
    UserForm11.Hide
    Clone2
End Sub

Private Sub Image8_Click()
    UserForm11.Hide
    Clone20
End Sub

Private Sub Image9_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image9_Click()
On Error Resume Next
    UserForm11.Hide
    LabourIntensity_SP_Query
    Components_SP_Query_
    Обновить
    aligment.aligment
    Обновить
    Aligment4d
    Обновить
    SaveToEXL
'    SaveToPDF
'    Python
'    Обновить
End Sub
