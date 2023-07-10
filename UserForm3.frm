VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Опция"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
UserForm3.Hide
beging:
 Start = Now()
 Dim FilesToOpen
 Dim ThisWorkbook, importWB As Workbook
 Dim SheetName As String
 Dim ws As Worksheet
 
 Set ThisWorkbook = ActiveWorkbook
 On Error GoTo ExitHandler
 SheetName = "ССЧ21"
 
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False

ThisWorkbook.Sheets(SheetName).Activate

 FilesToOpen = Application.GetOpenFilename _
 (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
 MultiSelect:=True, Title:="Выберите файл с численностью и текучестью кадров за 2021 год")

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
'X = ThisWorkbook.Sheets(SheetName).Range("AG5").Value2
'Y = ThisWorkbook.Sheets("Calculation21").Range("E2").Value2
'If X <> Y Then
'    MsgBox "Внимание!" _
'    & vbCr & "Загруженные данные по численности сотрудников не совпадают " _
'    & "с типом выбранной компании в расчётной ведомости!" _
'    & vbCr & "Численость рассчитана не корректно!" _
'    , vbCritical
'    result = MsgBox("Загрузить корректную расчётную ведомость?", vbYesNo)
'    If result = vbYes Then
'        Application.Run "Data_insertion"
'    Else: MsgBox "Действие отменено!" _
'    & vbCr & "Выберите корректный отчёт по численности с компанией " _
'    & vbCr & ThisWorkbook.Sheets("Calculation21").Range("E2").Value2
'
'    End If
'    GoTo beging
'End If
MsgBoxEx "Данные c численностью по компании" _
& vbCr & ThisWorkbook.Sheets(SheetName).Range("AI5").Value2 _
& vbCr & "за 2021 год" _
& vbCr & "добавлены успешно", 0, "Выполнено", 20
'MsgBoxEx "Выполнено 5%", 0, "5%. Мы только начали...", 5
ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub

Private Sub CommandButton2_Click()
UserForm3.Hide
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
 MultiSelect:=True, Title:="Выберите файл с численностью и текучестью кадров за 2022 год")

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
''наименование компании
'X = ThisWorkbook.Sheets(SheetName).Range("AI5").Value2
''проверка сумм
'Y = ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
'If X <> Y Then
'    MsgBox "Внимание!" _
'    & vbCr & "Загруженные данные по численности сотрудников не совпадают " _
'    & "с типом выбранной компании в расчётной ведомости!" _
'    & vbCr & "Численость рассчитана не корректно!" _
'    , vbCritical
'    result = MsgBox("Загрузить корректную расчётную ведомость?", vbYesNo)
'    If result = vbYes Then
'        Application.Run "Data_insertion"
'    Else: MsgBox "Действие отменено!" _
'    & vbCr & "Выберите корректный отчёт по численности с компанией " _
'    & vbCr & ThisWorkbook.Sheets("Calculation22").Range("E2").Value2
'
'    End If
'    GoTo beging
'End If
If ThisWorkbook.Sheets(SheetName).Range("AI2").Value2 = True Then
    MsgBoxEx "Данные с численностью по компании" _
    & vbCr & ThisWorkbook.Sheets(SheetName).Range("AI5").Value2 _
    & vbCr & "за 2022 год" _
    & vbCr & "добавлены успешно", 0, "Выполнено", 20
'ElseIf ThisWorkbook.Sheets(SheetName).Range("AI2").Value2 = alse Then
'    MsgBox "В загруженных данных обнаружены ошибки"
End If

ExitHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
 ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub

