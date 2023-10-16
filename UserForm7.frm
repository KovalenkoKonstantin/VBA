VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "Опции получения данных"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton16_Click()
    UserForm7.Hide
    LabourIntensity_SP_Query
End Sub

Private Sub CommandButton17_Click()
    UserForm7.Hide
    Components_SP_Query_
End Sub

Private Sub CommandButton18_Click()
    UserForm7.Hide
    GetExpendituresRefresh_SP_Query
End Sub

Private Sub CommandButton19_Click()
    UserForm7.Hide
    GetOrganizationRefresh_SP_Query
End Sub

Private Sub CommandButton20_Click()
    UserForm7.Hide
    GetProjectRefresh_SP_Query
End Sub

Private Sub CommandButton21_Click()
    UserForm7.Hide
    GetGozAttributeRefresh_SP_Query
End Sub
