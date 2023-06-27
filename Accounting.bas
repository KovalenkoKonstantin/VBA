Attribute VB_Name = "Accounting"
Sub Data_insertion_26()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Ан.сч26"
 awLastCol = 9
 Start = Now()
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
 MultiSelect:=True, Title:="Выберите файл с анализом 26 счёта")

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
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "D").End(xlUp).row
 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    
'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
MsgBox ("Данные успешно вставлены")

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets(SheetName).Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Sub Data_insertion_44()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "Ан.сч44"
 awLastCol = 9
 Start = Now()
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
 MultiSelect:=True, Title:="Выберите файл с анализом 44 счёта")

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
awLastRow = Cells(Rows.Count, "D").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "D").End(xlUp).row
 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    
'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
MsgBox ("Данные успешно вставлены")

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets(SheetName).Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

Sub Data_insertion_OFR()
 
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "ОФР"
 awLastCol = 20
 Start = Now()
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xls), *.xls", _
 MultiSelect:=True, Title:="Выберите файл с первым листом ОФР")

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
awLastRow = Cells(Rows.Count, "N").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column

Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 'вставка данных
 importWB.Sheets(1).Activate
 iwLastRow = Cells(Rows.Count, "N").End(xlUp).row
 importWB.Activate
 Cells.Select
    Selection.NumberFormat = _
        "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
    
'добавление формул
If ThisWorkbook.Sheets(SheetName).Range("X1").Value2 = 2021 Then
    'очистка
    ThisWorkbook.Sheets(SheetName).Range("AE2:AF6").Select
    Selection.ClearContents
    '2020
    ThisWorkbook.Sheets(SheetName).Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C28,C21,RC29)"
    Range("AE2").Select
    Selection.AutoFill Destination:=Range("AE2:AE7"), Type:=xlFillDefault
    Range("AE2:AF7").Select
    Range("AE2:AF7").Copy
    Range("AE2:AF7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '2021
    ThisWorkbook.Sheets(SheetName).Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C25,C21,RC29)"
    Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF7"), Type:=xlFillDefault
    Range("AF2:AF7").Select
    Range("AF2:AF7").Copy
    Range("AF2:AF7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ElseIf ThisWorkbook.Sheets(SheetName).Range("X1").Value2 = 2022 Then
    'очистка
    ThisWorkbook.Sheets(SheetName).Range("AF2:AG7").Select
    Selection.ClearContents
    '2021
    ThisWorkbook.Sheets(SheetName).Range("AF2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C28,C21,RC29)"
    Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF7"), Type:=xlFillDefault
    Range("AF2:AF7").Select
    Range("AF2:AF7").Copy
    Range("AF2:AF7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '2022
    ThisWorkbook.Sheets(SheetName).Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C25,C21,RC29)"
    Range("AG2").Select
    Selection.AutoFill Destination:=Range("AG2:AG7"), Type:=xlFillDefault
    Range("AG2:AG7").Select
    Range("AG2:AG7").Copy
    Range("AG2:AG7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End If
    
'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
Company = ThisWorkbook.Sheets(SheetName).Cells(10, 22).Value2
Period = ThisWorkbook.Sheets(SheetName).Cells(10, 23).Value2
'MsgBox ("Данные по компании " & Company & Period & " успешно добавлены")

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
' ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub

'Sub auto_open()                     'Макрос для автоматического запуска нужного макроса
'    Application.Run ("OpenUserForm")
'End Sub

Sub OpenUserForm()              'Нужный макрос
    UserForm1.Show                'Показать форму
End Sub

Sub Data_insertion_SS4()
beging:
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "ССЧ22"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл с численностью и текучестью кадров")

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
awLastRow = Cells(Rows.Count, "A").End(xlUp).row
'awLastCol = Cells(begin, Columns.Count).End(xlUp).Column
awLastCol = 29
Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
 With Selection
        .Clear
 End With

 'вставка данных
For Each ws In importWB.Sheets
 ws.Activate
 iwLastRow = Cells(Rows.Count, "A").End(xlUp).row

 importWB.Activate
 Range(Cells(1, 1), Cells(iwLastRow, awLastCol)).Copy
 
 ThisWorkbook.Sheets(SheetName).Activate
 awFirstRow = 1
 awFirstRow = Cells(Rows.Count, "A").End(xlUp).row
 awFirstCol = 1
 
 Range(Cells(awFirstRow + 1, awFirstCol), Cells(iwLastRow + awFirstRow, awLastCol)).Select
    With Selection
           .PasteSpecial Paste:=xlPasteAll
           .UnMerge
           .Font.Name = "Times New Roman"
           .WrapText = False
           .MergeCells = False
           .Font.Size = 8
    End With
Next ws
'завершение
importWB.Close
ThisWorkbook.Sheets(SheetName).Activate
X = ThisWorkbook.Sheets(SheetName).Range("AG5").Value2
Y = ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
If X <> Y Then
    MsgBox "Внимание!" _
    & vbCr & "Загруженные данные по численности сотрудников не совпадают " _
    & "с типом выбранной компании в расчётной ведомости!" _
    & vbCr & "Численость рассчитана не корректно!" _
    , vbCritical
    result = MsgBox("Загрузить корректную расчётную ведомость?", vbYesNo)
    If result = vbYes Then
        Application.Run "Data_insertion"
    Else: MsgBox "Действие отменено!" _
        & vbCr & "Выберите корректный отчёт по численности с компанией " _
        & vbCr & ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
    End If
    GoTo beging
End If
MsgBox "Данные по численности добавлены"

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
' ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler

End Sub
