VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
 Dim ThisWorkbook As Workbook
 Dim st As String

 Set ThisWorkbook = ActiveWorkbook

 On Error GoTo ExitHandler
 Application.Calculation = xlManual
 
' Application.ScreenUpdating = False
' Application.EnableEvents = False
' ActiveSheet.DisplayPageBreaks = False
' Application.DisplayStatusBar = False
' Application.DisplayAlerts = False
 
  ThisWorkbook.Sheets("�5").Activate
' ���� ����� ������� ��
For i = 1 To 50
    If Worksheets("�5").Cells(3, i) = "�������� ������� ����� �������� �����" Then
        column = i
    End If
Next

If Worksheets("�5").Cells(4, column) = False Then
    GoTo First
ElseIf Worksheets("�5").Cells(4, column) = True Then
    GoTo Second
End If

' ���� ���������� ��������
First:
For i = 5 To 500
    If Not Worksheets("�5").Cells(i, column) = "" And Worksheets("�5").Cells(i, column) = False Then
        Cells(i, column).Select
        ActiveCell.EntireRow.Select
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        Cells(i, column).Select
'        range(Cells(i, column), Cells(i, column)).Activate
        ActiveCell.FormulaR1C1 = "=CountByColor(RC[-25]:RC[-8],R2C16)=0"
    End If
Next
GoTo ExitHandler

Second:
For i = 5 To 500
    If Not Worksheets("�5").Cells(i, column) = "" And Worksheets("�5").Cells(i, column) = True Then
        Cells(i, column).Select
        ActiveCell.EntireRow.Select
        With Selection.Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
        Cells(i, column).Select
'        Cells(i, column).Activate
        ActiveCell.FormulaR1C1 = "=CountByColor(RC[-25]:RC[-8],R2C16)=0"
    End If
Next
    
ExitHandler:
       Application.ScreenUpdating = True
       Application.EnableEvents = True
       ActiveSheet.DisplayPageBreaks = True
       Application.DisplayStatusBar = True
       Application.DisplayAlerts = True
       Application.Calculation = xlAutomatic
       [t1].Activate
    Exit Sub
End Sub





