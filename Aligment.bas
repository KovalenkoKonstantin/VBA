Attribute VB_Name = "Aligment"
Sub aligment()
Start = Now()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

'Application.StatusBar = "��������� ����������� " & (Now() - Start) * 24 * 60 * 60 & " ������"

On Error Resume Next

I = 1
J = "variable2"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������2

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    I = "K"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData + 1) <> 0 Then
        Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "N"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData + 1) <> 0 Then
        Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "G"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData + 1) <> 0 Then
        Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "Q"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData + 1) <> 0 Then
        Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    
Application.StatusBar = "��������� 30%"
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "�8"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
ThisWorkbook.Sheets(sh).Range("F" & RowData & ":" & "O" & RowData).ClearContents

    I = "F"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData - 2) <> 0 Then
        Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "I"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData - 2) <> 0 Then
        Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "L"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData - 2) <> 0 Then
        Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "O"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData - 2) <> 0 Then
        Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    
Application.StatusBar = "��������� 60%"
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    I = "L"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData + 2) <> 0 Then
        Range(I & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    I = "O"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    If Range(I & RowData + 2) <> 0 Then
        Range(I & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    End If
    
Application.StatusBar = "��������� 90%"
    
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub

Sub Aligment��()
Start = Now()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

On Error Resume Next
    
'_________________________________________________________________________________

'I = 1
'J = "variable"
'sh = "��.��"
'ThisWorkbook.Sheets(sh).Activate
'RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������
'
''ThisWorkbook.Sheets(sh).Activate
'
''������� �������������
'ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "D" & RowData).ClearContents
'
'    I = "D"
'    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
'    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "11. ��"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
'ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "D" & RowData).ClearContents

    I = "C"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
    I = "F"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
    I = "I"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
    
    
    
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub
