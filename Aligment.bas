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

i = 1
j = "variable2"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row '��� ��������2

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    i = "K"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData + 1) <> 0 Then
        Range(i & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "N"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData + 1) <> 0 Then
        Range(i & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "G"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData + 1) <> 0 Then
        Range(i & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "Q"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData + 1) <> 0 Then
        Range(i & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    
Application.StatusBar = "��������� 30%"
    
'_________________________________________________________________________________

i = 1
j = "variable"
sh = "�8"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row '��� ��������

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
ThisWorkbook.Sheets(sh).Range("F" & RowData & ":" & "O" & RowData).ClearContents

    i = "F"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData - 2) <> 0 Then
        Range(i & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "I"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData - 2) <> 0 Then
        Range(i & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "L"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData - 2) <> 0 Then
        Range(i & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "O"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData - 2) <> 0 Then
        Range(i & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    
Application.StatusBar = "��������� 60%"
    
'_________________________________________________________________________________

i = 1
j = "variable"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row '��� ��������

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    i = "L"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData + 2) <> 0 Then
        Range(i & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    End If
    i = "O"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    If Range(i & RowData + 2) <> 0 Then
        Range(i & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
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

i = 1
j = "variable"
sh = "11. ��"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row '��� ��������

'ThisWorkbook.Sheets(sh).Activate

'������� �������������
'ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "D" & RowData).ClearContents

    i = "C"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    Range(i & RowData - 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    
    i = "F"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    Range(i & RowData - 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    
    i = "I"
    ThisWorkbook.Sheets(sh).Range(i & RowData) = 0
    Range(i & RowData - 1).GoalSeek Goal:=0, ChangingCell:=Range(i & RowData)
    
    
    
    
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub

Sub Aligment4d()
Start = Now()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

On Error Resume Next

i = 34
j = "variable"
sh = "4�"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row '��� ��������

    i = "AK"
    K = "S"
    ThisWorkbook.Sheets(sh).Range(i & RowData + 6) = 0
    Range(i & RowData + 5).GoalSeek Goal:=0, ChangingCell:=Range(K & RowData)
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub
