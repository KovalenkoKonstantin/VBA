Attribute VB_Name = "Aligment"
Sub Aligment12toP8()
Start = Now()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook
Dim L As Integer
Dim M As Integer

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

'�������� �� ��� ������
    For L = 1 To 20
        Worksheets(sh).Outline.ShowLevels rowLevels:=L
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next L
    For M = 1 To 20
        Worksheets(sh).Outline.ShowLevels columnLevels:=M
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next M
sh = "�8"
ThisWorkbook.Sheets(sh).Activate
    For L = 1 To 20
        Worksheets(sh).Outline.ShowLevels rowLevels:=L
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next L
    For M = 1 To 20
        Worksheets(sh).Outline.ShowLevels columnLevels:=M
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next M

sh = "12"
ThisWorkbook.Sheets(sh).Activate
    
'����������
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������2

'������� �������������
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    I = "K"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "N"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "G"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "Q"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
''������ �� ��� ������
'Worksheets(sh).Outline.ShowLevels 1, 1
'���������
Application.StatusBar = "��������� 30%"
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "�8"
ThisWorkbook.Sheets(sh).Activate
'����������
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������

'������� �������������
ThisWorkbook.Sheets(sh).Range("F" & RowData & ":" & "O" & RowData).ClearContents

    I = "F"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "I"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "L"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "O"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData - 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
Application.StatusBar = "��������� 60%"
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
'����������
RowData = Columns(I).Find(J, LookIn:=xlValues).row '��� ��������

'������� �������������
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    I = "L"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "O"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
'���������
Application.StatusBar = "��������� 90%"

'������ �� ��� ������
sh = "�8"
ThisWorkbook.Sheets(sh).Activate
Worksheets(sh).Outline.ShowLevels 1, 1
sh = "12"
ThisWorkbook.Sheets(sh).Activate
Worksheets(sh).Outline.ShowLevels 1, 1
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub
