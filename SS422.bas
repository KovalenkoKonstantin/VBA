Attribute VB_Name = "SS422"
Sub Insertion_ССЧ22()

 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "ССЧ22"
 Limit = 4 'последняя колонка базы
 begin = 15 'первый ряд вставки
 
 Dim aw(1 To 4) As Variant
 Dim iw(1 To 4) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.UnProtect Password:="123"
ThisWorkbook.Sheets(SheetName).Visible = True
ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл с численностью и текучестью кадров за год предшествующий предыдущему")

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

ThisWorkbook.Activate

'определение колонок рабочей книги
On Error Resume Next
For i = 1 To 15
    If Worksheets(SheetName).Cells(i, 1) = "Сотрудник" Then
        DataRow = i
    End If
Next i

For i = 1 To Limit
    If Worksheets(SheetName).Cells(DataRow, i) = "Сотрудник" Then
        aw(1) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Способ отражения" Then
        aw(2) = i
    End If
    If Worksheets(SheetName).Cells(DataRow, i) = "Списочн. численн." Then
        aw(3) = i
    End If
    If Worksheets(SheetName).Cells(DataRow + 1, i) = "Списочн. состава" Then
        aw(4) = i
    End If
Next i
 
 importWB.Sheets(1).Activate

'определение колонок импортируемой книги
For i = 1 To 30
    If importWB.Sheets(1).Cells(i, 1) = "Сотрудник" Then
        ImportFirstDataRow = i
    End If
Next i

For i = 1 To 30
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Сотрудник" Then '-
        iw(1) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Способ отражения" Then
        iw(2) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, i) = "Списочн. численн." Then
        iw(3) = i
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow + 1, i) = "Списочн. состава" Then
        iw(4) = i
    End If
Next i

'удаление предыдущих данных
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow + 3, Limit)).Select
 With Selection
        .Clear
 End With

 'статус бар
Application.StatusBar = "Вставка данных"

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row
 Range(Cells(1, 1), Cells(iwLastRow, 30)).Select
 With Selection
    .UnMerge
 End With

For i = 1 To 4
''статус бар
'Application.StatusBar = "Промежуточный цикл. Выполнено: " & Int(100 * i / Limit) & "%." & _
'" Общий прогресс: " & Int(87 * i / Limit) & "%" & _
'" Расчётное время до конца выполнения программы: " & _
'Int((100 - Int(87 * i / Limit)) * (((Now() - Start) * 24 * 60 * 60) / (Int(87 * i / Limit)))) & " секунд"
 'код вставки
 importWB.Activate
 Range(Cells(begin - 1, iw(i)), Cells(iwLastRow, iw(i))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(i)), Cells(iwLastRow, aw(i))).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next i

ThisWorkbook.Sheets(SheetName).Range("A2") = importWB.Sheets(1).Range("C4").Value2

'завершение
importWB.Close

ThisWorkbook.Sheets(SheetName).Protect Password:="123"
ThisWorkbook.Sheets(SheetName).Visible = False

MsgBoxEx "Данные c численностью по компании" _
& vbCr & ThisWorkbook.Sheets(SheetName).Range("A2").Value2 _
& vbCr & "добавлены успешно", 0, "Выполнено", 20

ExitHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ThisWorkbook.Sheets("Preferences").Activate
    ThisWorkbook.Sheets("Preferences").UnProtect Password:="123"
    Rows("81:91").EntireRow.AutoFit
    ThisWorkbook.Sheets("Preferences").Protect Password:="123"
    ThisWorkbook.Protect Password:="123"

 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub




