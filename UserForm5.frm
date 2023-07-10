VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Выберите вид расчётной ведомости"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ToggleButton2_Click()
    UserForm5.Hide
    UserForm2.Show
End Sub

Private Sub ToggleButton3_Click()
    UserForm5.Hide
    Project_Payroll_Insertion
End Sub

Private Sub ToggleButton4_Click()
    UserForm5.Hide
    UserForm6.Show
End Sub
