Attribute VB_Name = "Get_Data"
Sub Insertion()
 
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 ps = "123$"
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Temp"
 Limit = 2 'последняя колонка базы
 begin = 1 'первый ряд вставки

 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
'Application.Calculation = xlManual

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл")
 

 If TypeName(FilesToOpen) = "Boolean" Then ',была нажата кнопка отмены выход из процедуры
 GoTo ExitHandler
 End If

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1), UpdateLinks:=0)

On Error Resume Next

'ThisWorkbook.UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Visible = True
'ThisWorkbook.Sheets(SheetName).UnProtect Password:=ps
ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

ThisWorkbook.Activate

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

For i = 1 To Limit
 importWB.Activate
 Range(Cells(begin, i), Cells(iwLastRow - 1, i)).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, i), Cells(iwLastRow - 1, i)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With

Next i

'форматы
ThisWorkbook.Sheets(SheetName).Activate
Range(Cells(begin, 1), Cells(iwLastRow - 1, i)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(begin, 1), Cells(iwLastRow - 1, i)), , xlYes).Name = _
        "Sorce"
    ActiveSheet.ListObjects("Source").TableStyle = "TableStyleLight12"

'завершение
ThisWorkbook.Sheets(SheetName).Activate
'ThisWorkbook.Sheets(SheetName).Protect Password:=ps
'ThisWorkbook.Sheets(SheetName).Visible = False


ExitHandler:
    On Error Resume Next
    importWB.Close
'    Application.StatusBar = False
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    ActiveSheet.DisplayPageBreaks = True
'    Application.DisplayStatusBar = True
'    Application.DisplayAlerts = True
'    Application.Calculation = xlAutomatic
'    ThisWorkbook.Activate
'    ThisWorkbook.UnProtect Password:=ps
'    ThisWorkbook.Protect Password:=ps
'    ThisWorkbook.Protect Password:=ps
 Exit Sub
  
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub
