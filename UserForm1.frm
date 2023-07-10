VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Меню выбора действия"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

 UserForm1.Hide
 UserForm4.Show

End Sub



Private Sub CommandButton10_Click()
 UserForm1.Hide
 BudgetInsertion

' Dim FilesToOpen
' Dim ThisWorkbook, importWB As Workbook
' Dim SheetName As String
' Dim I As Worksheet
'
' Set ThisWorkbook = ActiveWorkbook
' On Error GoTo ErrHandler
' SheetName = "Бюджет"
' awLastCol = 37
' SearchRow = "A"
' UserMessage = "Ура!"
'
'Application.ScreenUpdating = False
'Application.EnableEvents = False
'ActiveSheet.DisplayPageBreaks = False
'Application.DisplayStatusBar = False
'Application.DisplayAlerts = False
'
'ThisWorkbook.Sheets(SheetName).Activate
'
' FilesToOpen = Application.GetOpenFilename _
' (FileFilter:="Microsoft Excel Files (*.xlsx), *.xlsx", _
' MultiSelect:=True, Title:="Выберите файл с бюджетом проекта по сотрудникам за весь период действия проекта")
'
' If TypeName(FilesToOpen) = "Boolean" Then
'    GoTo ExitHandler
' End If
'
'ThisWorkbook.Sheets(SheetName).Activate
'On Error Resume Next
'ActiveSheet.ShowAllData
'
''вставка листов
'Set importWB = Workbooks.Open(Filename:=FilesToOpen(1))
'
''удаление предыдущих данных
' ThisWorkbook.Sheets(SheetName).Activate
'awLastRow = Cells(Rows.Count, SearchRow).End(xlUp).row
'
'
'Range(Cells(1, 1), Cells(awLastRow, awLastCol)).Select
' With Selection
'        .Clear
' End With
'
' 'вставка данных
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
'
''завершение
'importWB.Close
'
'MsgBoxEx "Данные добавлены", 0, "Выполнено", 5
'
'ExitHandler:
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'    ActiveSheet.DisplayPageBreaks = True
'    Application.DisplayStatusBar = True
'    Application.DisplayAlerts = True
'    ThisWorkbook.Sheets("Preferences").Activate
' Exit Sub
'
'ErrHandler:
' MsgBox Err.Description
' Resume ExitHandler
End Sub

Private Sub CommandButton11_Click()
UserForm1.Hide
EmployeesInsertion
End Sub

Private Sub CommandButton12_Click()
    UserForm1.Hide
    UserForm6.Show
End Sub

Private Sub CommandButton3_Click()
 UserForm1.Hide
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
 MsgBox ("Действие отменено")
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
MsgBoxEx ("Отчёт о финансовых результатах по компании " & Company & Period & " успешно добавлен"), 0, "Выполнено", 15

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

Private Sub CommandButton4_Click()
    UserForm1.Hide
    UserForm5.Show
End Sub

Private Sub CommandButton5_Click()
    UserForm1.Hide
    UserForm3.Show
End Sub

Private Sub CommandButton6_Click()
    UserForm1.Hide
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
 ThisWorkbook.Sheets("Preferences").Activate
' Finish = (Now() - Start) * 24 * 60 * 60
'MsgBox (Finish)
 Exit Sub
 
ErrHandler:
 MsgBox Err.Description
 Resume ExitHandler
End Sub

Private Sub CommandButton7_Click()
    UserForm1.Hide
End Sub

Private Sub CommandButton8_Click()
UserForm1.Hide
Project_Payroll_Insertion
End Sub

Private Sub CommandButton9_Click()
    UserForm1.Hide
End Sub

'Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'    UserForm1.Hide
'    SaveToPDF
'End Sub

Private Sub Image1_Click()
 UserForm1.Hide
 SaveToPDF

MsgBox ("Файл сохранён в формате PDF в корневой папке")

End Sub

Private Sub Image3_click()
 
 UserForm1.Hide
 LayoutOff
 
End Sub

Private Sub Image4_Click()
 
 UserForm1.Hide
 LayoutOn
 
End Sub

Private Sub Label1_Click()
    UserForm1.Hide
 SaveToPDF

MsgBox ("Файл сохранён в формате PDF в корневой папке")
End Sub
