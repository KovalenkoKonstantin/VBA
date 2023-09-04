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

'Application.StatusBar = "Программа выполняется " & (Now() - Start) * 24 * 60 * 60 & " секунд"

On Error Resume Next

i = 1
j = "variable2"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row 'ряд значения2

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
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
    
Application.StatusBar = "Выполнено 30%"
    
'_________________________________________________________________________________

i = 1
j = "variable"
sh = "П8"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row 'ряд значения

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
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
    
Application.StatusBar = "Выполнено 60%"
    
'_________________________________________________________________________________

i = 1
j = "variable"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row 'ряд значения

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
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
    
Application.StatusBar = "Выполнено 90%"
    
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub

Sub AligmentПМ()
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
'sh = "ПМ.НР"
'ThisWorkbook.Sheets(sh).Activate
'RowData = Columns(I).Find(J, LookIn:=xlValues).row 'ряд значения
'
''ThisWorkbook.Sheets(sh).Activate
'
''очистка коэффициентов
'ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "D" & RowData).ClearContents
'
'    I = "D"
'    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
'    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
'_________________________________________________________________________________

i = 1
j = "variable"
sh = "11. НР"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row 'ряд значения

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
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
sh = "4д"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(i).Find(j, LookIn:=xlValues).row 'ряд значения

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
