Attribute VB_Name = "Support"
Sub word()
MsgBox "отключено"
End Sub
Sub CleanIt()

Dim row, column, X As Integer
On Error GoTo ErrHandler
 Application.ScreenUpdating = False
 Application.EnableEvents = False
 ActiveSheet.DisplayPageBreaks = False
 Application.DisplayStatusBar = False
 Application.DisplayAlerts = False

' ищем номер строки по "Количество этапов"
For i = 1 To 50
    If Worksheets("Соисполнитель").Cells(i, 1) = "Количество этапов" Then
        row = i
    End If
Next

' номер колонки
column = 1


'If Application.Worksheets("Соисполнитель").Cells(row, column + 1).Value Is Empty Then
'    GoTo ErrHandler
'End If

X = Application.Worksheets("Соисполнитель").Cells(row, column + 1).Value 'значение ячейки

'удаляем лишние строки
For i = row + X - 1 To row + 1 Step -1
    Rows(i).EntireRow.delete
'    Range(i, column).EntireRow.Delete
Next

'очищаем значения ячеек
For i = 2 To 3
    Application.Worksheets("Соисполнитель").Cells(row, i).Clear
Next

ErrHandler:
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 ActiveSheet.DisplayPageBreaks = True
 Application.DisplayStatusBar = True
 Application.DisplayAlerts = True
End Sub


Sub Social_contribution()
 
' Dim FilesToOpen
 Dim ThisWorkbook As Workbook
' Dim ws, this As Worksheet
' Dim pt As PivotTable
' Dim с, d As Range
 Dim temp, temp1, temp2 As String
' Dim x As Integer
 X = "7,8"
 Y = "РВ2"
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 'поиск границ таблицы
 ThisWorkbook.Sheets(Y).Activate
 s = Cells(Rows.Count, "B").End(xlUp).row 'последний столбец данных
 K = Cells(2, Columns.Count).End(xlToLeft).column 'последняя колонка данных
 
 ' ищем номера колонок
For i = 1 To K
    'определяем колонку заработной платы
    If Worksheets(Y).Cells(2, i) = "Оклады" Then
        zp = i
    End If
    'определяем колонку социальных выплат
    If Worksheets(Y).Cells(2, i) = "% Страховых взносов" Then
        sp = i
    End If
    'определяем колонку года
    If Worksheets(Y).Cells(2, i) = "План" Then
        yr = i
    End If
    'определяем колонку проверки
    If Worksheets(Y).Cells(2, i) = "Проверка" Then
        check = i
    End If
Next i

'удаляем предыдущие значения
    ThisWorkbook.Sheets(X).Activate
    Range(check & "4:" & check & K).Clear
'main
For i = 3 To s
'skip constant rows
    If Worksheets(Y).Cells(i, 2).Value2 = "Итого" Then
        GoTo ExitHandler
    End If
    If Worksheets(Y).Cells(i, 2).Value2 = Worksheets(Y).Cells(i, 3).Value2 Then
        i = i + 1
    End If
'переносим значение зарплаты в лист расчётов
    Worksheets(Y).Cells(i, zp).Copy
    ThisWorkbook.Sheets(X).Activate
    Range("B4").Select
    With Selection
        .PasteSpecial Paste:=xlPasteValues
    End With
    
'меняем значения ячеек в листе с вычислениями социальных выплат
    temp = Worksheets(Y).Cells(i, yr).Value2
    ThisWorkbook.Sheets(X).Activate
    Range("J4:J15").Clear
        For C = 4 To 15
            temp1 = Cells(C, 1).Value2
            Cells(C, 10) = temp1 & " " & temp
        Next C
 'ищем совпадение месяцев
    For j = 4 To 15
'        a = ThisWorkbook.Sheets(x).Cells(j, 10).Value2
'        b = Worksheets(y).Cells(i, 3).Value2
        If ThisWorkbook.Sheets(X).Cells(j, 10).Value2 = Worksheets(Y).Cells(i, 3).Value2 Then
            destinct = j 'нужный ряд для переноса значения в расчётную ведомость
        End If
    Next j
'удаляем внесённые ключи из 7,8
    ThisWorkbook.Sheets(X).Activate
    Range("J4:J15").Clear
 'переносим % соц выплат в расчётную ведомость
 ThisWorkbook.Sheets(X).Activate
 Cells(destinct, 9).Copy
 ThisWorkbook.Sheets(Y).Activate
 Cells(i, check).Select
    With Selection
        .PasteSpecial Paste:=xlPasteValues
    End With
Next i


ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub


Sub RefreshOrder()
Dim objWord As Object
Dim FileStart
Dim FileNew

Set objWord = CreateObject("Word.Application")

    FileSt = "D:\РКМ\ТФЦ\022-7\Приказ.docx"
    FileNew = "D:\РКМ\ТФЦ\022-7\Приказ1.docx"

    objWord.Documents.Open FileSt
                
    For Each MyLink In objWord.ActiveDocument.Fields
        MyLink.Update
        MyLink.Unlink
    Next MyLink

    objWord.ActiveDocument.SaveAs _
            Filename:=FileNew, _
            FileFormat:=wdFormatDocument, _
            Password:="", _
            AddToRecentFiles:=True, _
            WritePassword:="", _
            ReadOnlyRecommended:=False
objWord.Quit
End Sub

Sub Budget()

 Dim ThisWorkbook, importWB As Workbook
 Dim FilesToOpen
' Dim MyRange, MyCell As range
 Dim key As String
 Ye_ar = 2022 'начальный год бюджета
 X = 4 'количество листов для вставки
 DataTab = "Бюджет" 'лист данных
 WorkTab = "НЧ" 'рабочий лист
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Файл для вставки")
 
 If TypeName(FilesToOpen) = "Boolean" Then
 GoTo ExitHandler
 End If
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 
 'поиск границ таблицы данных
 ThisWorkbook.Sheets(DataTab).Activate
 FirstRowData = Columns(1).Find("*", LookIn:=xlValues).row 'ряд первого значения
 LastRowData = Cells(Rows.Count, 2).End(xlUp).row 'последний ряд данных
 LastColumnData = Cells(FirstRowData, Columns.Count).End(xlToLeft).column 'последняя колонка данных
  
 'ищем номера колонок в листе данных
For i = 1 To LastColumnData
    'определяем колонку ключа1
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Ключ1" Then
        Key1ColData = i
    End If
    'определяем колонку имени сотрудника
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Сотрудник" Then
        EmployeeColData = i
    End If
    'определяем колонку должности
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Должность" Then
        PositionColData = i
    End If
    'определяем колонку года
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Год" Then
        YearColData = i
    End If
    'определяем колонку Премии
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Премии" Then
        PrizeColData = i
    End If
    'определяем колонку ключа2
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Ключ2" Then
        Key2ColData = i
    End If
    'определяем колонку Второй должности
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Должность2" Then
        Position2ColData = i
    End If
    'определяем колонку месяца в числовом выражении
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Мес" Then
        MonthNumberColData = i
    End If
    'определяем колонку декабря
    If Worksheets(DataTab).Cells(FirstRowData, i) = "Декабрь" Then
        DecemberColData = i
    End If
Next i
 
 'удаление предыдущих данных
 Range(Cells(FirstRowData + 1, Key1ColData), Cells(LastRowData, LastColumnData)).Select
 With Selection
        .ClearContents
 End With

'вставка листов
Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
For i = 1 To X
    On Error Resume Next
     importWB.Sheets(i).Activate
     lLastRow = Cells(Rows.Count, 1).End(xlUp).row
     j = lLastRow
     
     importWB.Sheets(i).Activate
     Range("A3:N" & j).Select
     Range("A3:N" & j).Copy
     ThisWorkbook.Sheets(DataTab).Activate
     'ищем новое значение последнего ряда
     lLastRow = Cells(Rows.Count, 3).End(xlUp).row
     jRenow = lLastRow
     Range(Cells(jRenow + 1, EmployeeColData), Cells(jRenow + j, DecemberColData)).Select
     With Selection
            .PasteSpecial Paste:=xlPasteAll
            .UnMerge
            .Font.Name = "Times New Roman"
            .WrapText = False
            .MergeCells = False
            .Font.Size = 8
     End With
     
     'добавление идентификатора года в колонку премии
     lLastRow = Cells(Rows.Count, 2).End(xlUp).row
     jNew = lLastRow
     Range(Cells(jRenow + 1, PrizeColData), Cells(jNew, PrizeColData)).Value2 = i

Next i
'закрытие книги данных
importWB.Close

'вставка формулы расчёта года
 Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).FormulaR1C1 = _
    "=IF(RC[1]=1,2022,IF(RC[1]=2,2023,IF(RC[1]=3,2022,IF(RC[1]=4,2023))))"
    'вставка значений вместо формулы
    Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).Select
    Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).Copy
    Range(Cells(FirstRowData + 1, YearColData), Cells(jNew, YearColData)).Select
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
        .Font.Size = 8
 End With

'добавление идентификатора премии в колонку ключа2
    Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaR1C1 = "=IF(OR(RC[-1]=3,RC[-1]=4),""Премия"","""")"
    'перенос значений премий в колонку
    Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).Select
    Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).Copy
    Range(Cells(FirstRowData + 1, PrizeColData), Cells(jNew, PrizeColData)).Select
 With Selection
        .PasteSpecial Paste:=xlPasteValues
        .UnMerge
        .Font.Name = "Times New Roman"
        .WrapText = False
        .MergeCells = False
        .Font.Size = 8
 End With
 
 'добавление формулы ключа в первую колонку
Range(Cells(FirstRowData + 1, 1), Cells(jNew, 1)).FormulaR1C1 = "=CONCATENATE(RC[2],RC[16],RC[17])"

'добавление формулы ключа1
Range(Cells(FirstRowData + 1, Key1ColData), Cells(jNew, Key1ColData)).FormulaR1C1 = "=CONCATENATE(RC[1],RC[15],RC[2],RC[16])"

'добавление формулы ключа2
Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaR1C1 = "=RC[-18]"

'добавление формулы должности2
Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaR1C1 = _
        "=IF(AND(RC[-17]=R[1]C[-17],RC[-16]<>R[1]C[-16]),R[1]C[-16],"""")"
        
'добавление формулы месяца в числовом выражении для областей с вторыми должностями
Range(Cells(FirstRowData + 1, Key2ColData), Cells(jNew, Key2ColData)).FormulaArray = _
        "=IF(RC[-1]="""","""",MATCH(TRUE(),(RC[-16]:RC[-5]=""""),FALSE()))"
 
' 'удаление контента из колонки ключа2
' range(Cells(FirstRowData + 1, LastColumnData + 1), Cells(jNew, LastColumnData + 1)).ClearContents

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
'    ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Sub RefreshPivots()
Dim pt As PivotTable
Dim ws As Worksheet
Dim ThisWorkbook As Workbook

Set ThisWorkbook = ActiveWorkbook
On Error GoTo ExitHandler

Application.ScreenUpdating = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
Application.Calculation = xlManual

For Each ws In ThisWorkbook
pt.Refresh
Next ws

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
' ThisWorkbook.Sheets("Preferences").Activate
 Exit Sub
End Sub

Sub ShowTabs()
 Dim tb
 On Error Resume Next
 For Each tb In Worksheets
 tb.Visible = True
 Next
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub HideSys()
Application.Calculation = xlManual
 Dim ws As Worksheet
 On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value2 = "sys" _
        Or ws.Range("A1").Value2 = "Трудоёмкость" _
        Or ws.Range("A1").Value2 = "Статья затрат" Or ws.Range("A1").Value2 = "Имя" _
        Or ws.Range("A1").Value2 = "company_name" _
        Or ws.Range("A1").Value2 = "organization_id" _
        Or ws.Range("A1").Value2 = "Наименование статей в 1С" _
        Or ws.Range("H2").Value2 = "Отчет о финансовых результатах" _
        Or ws.Range("J1").Value2 = "Сумма" Then
            ws.Visible = False
        End If
    Next ws
Application.Calculation = xlAutomatic
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub UnhideSys()
 Dim ws As Worksheet
 On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value2 = "sys" _
        Or ws.Range("A1").Value2 = "Трудоёмкость" _
        Or ws.Range("A1").Value2 = "Статья затрат" Or ws.Range("A1").Value2 = "Имя" _
        Or ws.Range("A1").Value2 = "company_name" _
        Or ws.Range("A1").Value2 = "Наименование статей в 1С" _
        Or ws.Range("H2").Value2 = "Отчет о финансовых результатах" _
        Or ws.Range("J1").Value2 = "Сумма" Then
            ws.Visible = True
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub HideEmpty()
 Dim ws As Worksheet
 On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value2 = "1" Then
            ws.Visible = True
            ws.Select
            With ws.Tab
                .ColorIndex = xlNone
                .TintAndShade = 0
            End With
            ws.Visible = False
        End If
    Next ws
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub Protect()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:="123"
    Next ws
    ThisWorkbook.Protect Password:="123"
 ActiveWorkbook.Sheets("Preferences").Activate
End Sub

Sub UnProtect()
 Dim ws As Worksheet
 On Error GoTo errorhandler
    For Each ws In ThisWorkbook.Worksheets
        ws.UnProtect Password:="123"
    Next ws
    ThisWorkbook.UnProtect Password:="123"

errorhandler:
' MsgBox ("Все листы разблокированы")
ActiveWorkbook.Sheets("Preferences").Activate

End Sub

Sub button1()
    [J61] = 0
    Range("J60").GoalSeek Goal:=0, ChangingCell:=Range("J61")
End Sub
Sub button2()
    [J61] = 0
    Range("J60").GoalSeek Goal:=0, ChangingCell:=Range("J61")
End Sub

