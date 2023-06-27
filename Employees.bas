Attribute VB_Name = "Employees"
Sub EmployeesInsertion()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim I As Worksheet

 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ErrHandler
 SheetName = "Employees"
 awLastCol = 37
 SearchRow = "A"
 UserMessage = "Ура!"

Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл с Сотрудниками по выбранной организации")

 If TypeName(FilesToOpen) = "Boolean" Then
    GoTo ExitHandler
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row


Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 'вставка данных
 For Each I In importWB.Sheets
     I.Activate
     iwLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
     importWB.Activate
     Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
    
     ThisWorkbook.Sheets(SheetName).Activate
     awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
     Range(Cells(awLastRow, 1), Cells(iwLastRow + awLastRow - 1, awLastCol)).Select
        With Selection
               .PasteSpecial Paste:=xlPasteAll
               .UnMerge
               .Font.Name = "Times New Roman"
               .WrapText = False
               .MergeCells = False
               .Font.Size = 10
        End With
 Next I

'завершение
importWB.Close

MsgBoxEx "Данные добавлены", 0, "Выполнено", 3

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub

ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub



