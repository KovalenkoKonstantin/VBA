Attribute VB_Name = "Budget"
Sub BudgetInsertion()
 
'Dim I As Worksheet
' 'вставка данных с многих листов
' For Each I In importWB.Sheets
'     I.Activate
'     iwLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
'     importWB.Activate
'     Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
'
'     ThisWorkbook.Sheets(SheetName).Activate
'     awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
'     Range(Cells(awLastRow, 1), Cells(iwLastRow + awLastRow - 1, awLastCol)).Select
'        With Selection
'               .PasteSpecial Paste:=xlPasteAll
'               .UnMerge
'               .Font.Name = "Times New Roman"
'               .WrapText = False
'               .MergeCells = False
'               .Font.Size = 10
'        End With
' Next I

'завершение


Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim ws, this As Worksheet
 Dim MyRange, MyCell As Range
 Dim SheetName, CompanyName As String
  
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Бюджет"
 DistinctYear = "2021 - 2024"
 SearchRow = "A"
 Limit = 54 'последняя колонка листа
 begin = 5 'первый ряд вставки
 CompanyName = ThisWorkbook.Sheets("Preferences").Range("C7").Value2 'имя проекта
 
 Dim aw(1 To 54) As Variant
 Dim iw(1 To 54) As Variant
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
Application.DisplayAlerts = False

FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл с бюджетом по компании " & CompanyName & " за " & DistinctYear & " года")
 
 'статус бар
Application.StatusBar = "Анализ данных..."

 If TypeName(FilesToOpen) = "Boolean" Then ',была нажата кнопка отмены выход из процедуры
 GoTo ExitHandler
 End If

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))

On Error Resume Next

 'проверка правильности выбора данных
 importWB.Sheets(1).Activate
 Range("A3").Select
 ActiveCell.FormulaR1C1 = "=COUNTIF(R[3]C[9]:R[3]C[56],""<>"""""")"
 If Range("A3").Value2 <> 48 Or Range("A10").Value2 <> CompanyName Then
    Range("A3").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "Выбран ненеправильный файл." _
    & vbCr & "Процесс прерван.", vbCritical, "Bad Day", 20
    GoTo ExitHandler
 Else
 Range("A3").Select
    With Selection:
        .Clear
    End With
    MsgBoxEx "Выбран правильный файл с бюджетом" _
    & vbCr & "Продолжаем.", 0, "Succes", 5
 End If

ThisWorkbook.Sheets(SheetName).Activate
On Error Resume Next
ActiveSheet.ShowAllData

ThisWorkbook.Activate

'определение колонок рабочей книги
On Error Resume Next
For I = 1 To 10
    If Worksheets(SheetName).Cells(I, 1) = "Должность" Then
        DataRow = I
    End If
Next I

For I = 1 To Limit
'статус бар
Application.StatusBar = "Определение колонок рабочей книги: " & Int(100 * I / Limit) & "%."

    If Worksheets(SheetName).Cells(DataRow, I) = "Должность" Then
        aw(1) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Начисление" Then
        aw(2) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Организация" Then
        aw(3) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Сотрудник" Then
        aw(4) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Проект" Then
        aw(5) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "График работы" Then
        aw(6) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Январь 2021" Then
        aw(7) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Февраль 2021" Then
        aw(8) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Март 2021" Then
        aw(9) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Апрель 2021" Then
        aw(10) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Май 2021" Then
        aw(11) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июнь 2021" Then
        aw(12) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июль 2021" Then
        aw(13) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Август 2021" Then
        aw(14) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Сентябрь 2021" Then
        aw(15) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Октябрь 2021" Then
        aw(16) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Ноябрь 2021" Then
        aw(17) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Декабрь 2021" Then
        aw(18) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Январь 2022" Then
        aw(19) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Февраль 2022" Then
        aw(20) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Март 2022" Then
        aw(21) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Апрель 2022" Then
        aw(22) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Май 2022" Then
        aw(23) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июнь 2022" Then
        aw(24) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июль 2022" Then
        aw(25) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Август 2022" Then
        aw(26) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Сентябрь 2022" Then
        aw(27) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Октябрь 2022" Then
        aw(28) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Ноябрь 2022" Then
        aw(29) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Декабрь 2022" Then
        aw(30) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Январь 2023" Then
        aw(31) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Февраль 2023" Then
        aw(32) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Март 2023" Then
        aw(33) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Апрель 2023" Then
        aw(34) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Май 2023" Then
        aw(35) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июнь 2023" Then
        aw(36) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июль 2023" Then
        aw(37) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Август 2023" Then
        aw(38) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Сентябрь 2023" Then
        aw(39) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Октябрь 2023" Then
        aw(40) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Ноябрь 2023" Then
        aw(41) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Декабрь 2023" Then
        aw(42) = I
    End If
' ------------------------------------------------------------------------------------------
    If Worksheets(SheetName).Cells(DataRow, I) = "Январь 2024" Then
        aw(43) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Февраль 2024" Then
        aw(44) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Март 2024" Then
        aw(45) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Апрель 2024" Then
        aw(46) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Май 2024" Then
        aw(47) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июнь 2024" Then
        aw(48) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Июль 2024" Then
        aw(49) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Август 2024" Then
        aw(50) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Сентябрь 2024" Then
        aw(51) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Октябрь 2024" Then
        aw(52) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Ноябрь 2024" Then
        aw(53) = I
    End If
    If Worksheets(SheetName).Cells(DataRow, I) = "Декабрь 2024" Then
        aw(54) = I
    End If
    
Next I
 
 importWB.Sheets(1).Activate

'определение колонок импортируемой книги
For I = 1 To 20
    If importWB.Sheets(1).Cells(I, 1) = "Организация" Then
        ImportFirstDataRow = I
    End If
Next I
'For I = 1 To 20
'    If importWB.Sheets(1).Cells(I, 1) = "Сотрудник" Then
'        ImportSecondDataRow = I
'    End If
'Next I

For I = 1 To Limit + 20
Application.StatusBar = "Определение колонок импортируемой книги: " & Int(100 * I / (Limit + 20)) & "%."

    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Должность" Then
        iw(1) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Начисление" Then
        iw(2) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Организация" Then
        iw(3) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Сотрудник" Then
        iw(4) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Проект" Then
        iw(5) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "График работы" Then
        iw(6) = I
    End If
' ------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Январь 2021" Then
        iw(7) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Февраль 2021" Then
        iw(8) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Март 2021" Then
        iw(9) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Апрель 2021" Then
        iw(10) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Май 2021" Then
        iw(11) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июнь 2021" Then
        iw(12) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июль 2021" Then
        iw(13) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Август 2021" Then
        iw(14) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Сентябрь 2021" Then
        iw(15) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Октябрь 2021" Then
        iw(16) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Ноябрь 2021" Then
        iw(17) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Декабрь 2021" Then
        iw(18) = I
    End If
' ------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Январь 2022" Then
        iw(19) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Февраль 2022" Then
        iw(20) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Март 2022" Then
        iw(21) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Апрель 2022" Then
        iw(22) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Май 2022" Then
        iw(23) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июнь 2022" Then
        iw(24) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июль 2022" Then
        iw(25) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Август 2022" Then
        iw(26) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Сентябрь 2022" Then
        iw(27) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Октябрь 2022" Then
        iw(28) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Ноябрь 2022" Then
        iw(29) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Декабрь 2022" Then
        iw(30) = I
    End If
' ------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Январь 2023" Then
        iw(31) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Февраль 2023" Then
        iw(32) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Март 2023" Then
        iw(33) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Апрель 2023" Then
        iw(34) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Май 2023" Then
        iw(35) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июнь 2023" Then
        iw(36) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июль 2023" Then
        iw(37) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Август 2023" Then
        iw(38) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Сентябрь 2023" Then
        iw(39) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Октябрь 2023" Then
        iw(40) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Ноябрь 2023" Then
        iw(41) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Декабрь 2023" Then
        iw(42) = I
    End If
' ------------------------------------------------------------------------------------------
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Январь 2024" Then
        iw(43) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Февраль 2024" Then
        iw(44) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Март 2024" Then
        iw(45) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Апрель 2024" Then
        iw(46) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Май 2024" Then
        iw(47) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июнь 2024" Then
        iw(48) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Июль 2024" Then
        iw(49) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Август 2024" Then
        iw(50) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Сентябрь 2024" Then
        iw(51) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Октябрь 2024" Then
        iw(52) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Ноябрь 2024" Then
        iw(53) = I
    End If
    If importWB.Sheets(1).Cells(ImportFirstDataRow, I) = "Декабрь 2024" Then
        iw(54) = I
    End If

Next I

'удаление предыдущих данных
Application.StatusBar = "Удаление предыдущих данных."
 ThisWorkbook.Sheets(SheetName).Activate
awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
awLastCol = Cells(begin, Columns.Count).End(xlUp).column
Range(Cells(begin, 1), Cells(awLastRow, Limit)).Select
 With Selection
        .Clear
 End With

 'определение последнего ряда IWB
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row

For I = 1 To Limit

'статус бар
Application.StatusBar = "Добавление рядов данных: " & Int(100 * I / Limit) & "%."
 
 'добавление
 importWB.Activate
 Range(Cells(begin + 3, iw(I)), Cells(iwLastRow, iw(I))).Copy

 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(begin, aw(I)), Cells(iwLastRow - 3, aw(I))).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next I

''статус бар
'Application.StatusBar = "Выполнено: 95 %"

'форматы
ThisWorkbook.Sheets(SheetName).Activate
Columns("G:BB").Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
        
'завершение
ThisWorkbook.Sheets(SheetName).Activate
MsgBoxEx "Бюджет по компании" _
    & vbCr & ThisWorkbook.Sheets("Бюджет").Range("C10").Value2 _
    & vbCr & "добавлен успешно", 0, "Выполнено", 25

ExitHandler:
    On Error Resume Next
    importWB.Close
    Application.StatusBar = False
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
