Attribute VB_Name = "Aligment"
Sub Aligment()
Start = Now()
Dim ThisWorkbook As Workbook
Set ThisWorkbook = ActiveWorkbook

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

Application.StatusBar = "Программа выполняется " & (Now() - Start) * 24 * 60 * 60 & " секунд"

On Error Resume Next

I = 1
J = "variable2"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row 'ряд значения2

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
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
    
Application.StatusBar = "Выполнено 30%"
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "П8"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row 'ряд значения

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
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
    
Application.StatusBar = "Выполнено 60%"
    
'_________________________________________________________________________________

I = 1
J = "variable"
sh = "12"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row 'ряд значения

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "Q" & RowData).ClearContents

    I = "L"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    I = "O"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 2).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
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

I = 1
J = "variable"
sh = "ПМ.НР"
ThisWorkbook.Sheets(sh).Activate
RowData = Columns(I).Find(J, LookIn:=xlValues).row 'ряд значения

'ThisWorkbook.Sheets(sh).Activate

'очистка коэффициентов
ThisWorkbook.Sheets(sh).Range("D" & RowData & ":" & "D" & RowData).ClearContents

    I = "D"
    ThisWorkbook.Sheets(sh).Range(I & RowData) = 0
    Range(I & RowData + 1).GoalSeek Goal:=0, ChangingCell:=Range(I & RowData)
    
    
Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
End Sub
