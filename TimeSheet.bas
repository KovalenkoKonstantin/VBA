Attribute VB_Name = "TimeSheet"
Sub TimeSheet()
beging:
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Табель"
 awLastCol = 63
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл с табелем рабочего времени")

 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 importWB.Sheets(1).Activate

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "AC").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 'вставка данных
For Each ws In importWB.Sheets
 ws.Activate
 iwLastRow = Cells(Rows.Count, "AC").End(xlUp).row

 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 awFirstRow = 1
 awFirstRow = Cells(Rows.Count, "AC").End(xlUp).row
 awFirstCol = 1
 
 Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
    End With
Next ws
'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate

MsgBoxEx "Табель рабочего времени добавлен", 0, "Выполнено", 15

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
ThisWorkbook.Sheets(SheetName).Visible = False
 ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub

