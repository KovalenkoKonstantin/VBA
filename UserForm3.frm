VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "�����"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
UserForm3.Hide
beging:
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "���21"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ���� � ������������ � ���������� ������ �� 2021 ���")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column
awLastCol = 29
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 '������� ������
For Each ws In importWB.Sheets
 ws.Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 awFirstRow = 1
 awFirstRow = Cells(Rows.Count, "A").End(xlUp).row
 awFirstCol = 1
 
 Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next ws
'����������
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
'X = ThisWorkbook.Sheets(SheetName).Range("AG5").Value2
'Y = ThisWorkbook.Sheets("Calculation21").Range("E2").Value2
'If X <> Y Then
'    MsgBox "��������!" _
'    & vbCr & "����������� ������ �� ����������� ����������� �� ��������� " _
'    & "� ����� ��������� �������� � ��������� ���������!" _
'    & vbCr & "���������� ���������� �� ���������!" _
'    , vbCritical
'    result = MsgBox("��������� ���������� ��������� ���������?", vbYesNo)
'    If result = vbYes Then
'        Application.Run "Data_insertion"
'    Else: MsgBox "�������� ��������!" _
'    & vbCr & "�������� ���������� ����� �� ����������� � ��������� " _
'    & vbCr & ThisWorkbook.Sheets("Calculation21").Range("E2").Value2
'
'    End If
'    GoTo beging
'End If
MsgBoxEx "������ c ������������ �� ��������" _
& vbCr & ThisWorkbook.Sheets(SheetName).Range("AI5").Value2 _
& vbCr & "�� 2021 ���" _
& vbCr & "��������� �������", 0, "���������", 20
'MsgBoxEx "��������� 5%", 0, "5%. �� ������ ������...", 5
ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub

Private Sub CommandButton2_Click()
UserForm3.Hide
beging:
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "���22"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="�������� ���� � ������������ � ���������� ������ �� 2022 ���")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'������� ������
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'�������� ���������� ������
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column
awLastCol = 29
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 '������� ������
For Each ws In importWB.Sheets
 ws.Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 awFirstRow = 1
 awFirstRow = Cells(Rows.Count, "A").End(xlUp).row
 awFirstCol = 1
 
 Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next ws
'����������
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
''������������ ��������
'X = ThisWorkbook.Sheets(SheetName).Range("AI5").Value2
''�������� ����
'Y = ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
'If X <> Y Then
'    MsgBox "��������!" _
'    & vbCr & "����������� ������ �� ����������� ����������� �� ��������� " _
'    & "� ����� ��������� �������� � ��������� ���������!" _
'    & vbCr & "���������� ���������� �� ���������!" _
'    , vbCritical
'    result = MsgBox("��������� ���������� ��������� ���������?", vbYesNo)
'    If result = vbYes Then
'        Application.Run "Data_insertion"
'    Else: MsgBox "�������� ��������!" _
'    & vbCr & "�������� ���������� ����� �� ����������� � ��������� " _
'    & vbCr & ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
'
'    End If
'    GoTo beging
'End If
If ThisWorkbook.Sheets(SheetName).Range("AI2").Value2 = True Then
    MsgBoxEx "������ � ������������ �� ��������" _
    & vbCr & ThisWorkbook.Sheets(SheetName).Range("AI5").Value2 _
    & vbCr & "�� 2022 ���" _
    & vbCr & "��������� �������", 0, "���������", 20
'ElseIf ThisWorkbook.Sheets(SheetName).Range("AI2").Value2 = alse Then
'    MsgBox "� ����������� ������ ���������� ������"
End If

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub

